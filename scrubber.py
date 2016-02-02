# scrubber.py
# http://sites.google.com/site/pocketsense/
# fix known issues w/ OFX downloads
# Ver 1: rlc, 2010

# 05-Aug-2010*rlc
# - Added _scrubTime() function to fix NULL time stamps so that transactions record on the correct date
#   regardless of time zone.

# 28-Jan-2011*rlc
#   - Added _scrubDTSTART() to fix missing <DTEND> fields when a <DTSTART> exists.
#   - Recoded scrub routines to use regex substitutions

# 28-Aug-2012*rlc
#   - Added quietScrub option to sites.dat (suppresses scrub messages)

# 02-Feb-2013*rlc
#   - Bug fix in _scrubDTSTART
#   - Added scrub routine to verify Investment buy and sell transactions

# 17-Feb-2013*rlc
#   - Added scrub routine for CORRECTACTION and CORRECTFITID tags (not supported by Money)

# 06-Jul-2013*rlc
#   - Bug fix in _scrubDTSTART()

# 20-Feb-2014*rlc
#   - Bug fix in _scrubINVsign() for SELL transactions

import os, sys, re, datetime
import site_cfg
from control2 import *

userdat = site_cfg.site_cfg()
nullTimeUpdated = False
stat=False

def scrubPrint(line):
    if not userdat.quietScrub:
        print line
    
def scrub(filename, site):
    #filename = string
    #site = DICT structure containing full site info from sites.dat
 
    siteURL = FieldVal(site, 'url').upper()
    dtHrs = FieldVal(site, 'timeOffset')
    f = open(filename,'r')
    ofx = f.read()  #as-found ofx message
    #print ofx
    
    if 'DISCOVERCARD' in siteURL: ofx= _scrubDiscover(ofx)
        
    ofx= _scrubTime(ofx)     #fix 000000 and NULL datetime stamps 

    if dtHrs <> 0: ofx = _scrubShiftTime(ofx, dtHrs)   #note: always call *after* _scrubTime()
    
    ofx= _scrubDTSTART(ofx)  #fix missing <DTEND> fields
      
    #fix malformed investment buy/sell signs (neg vs pos), if they exist
    if "<INVSTMTTRNRS>" in ofx.upper(): ofx= _scrubINVsign(ofx)  
        
    #perform general ofx cleanup
    ofx = _scrubGeneral(ofx)
    
    #close the input file
    f.close()
    
    #write the new version to the same file name
    f = open(filename, 'w')
    f.write(ofx)
    f.close    

#-----------------------------------------------------------------------------
# OFX.DISCOVERCARD.COM
#   1.  Discover OFX files will contain transaction identifiers w/ the following format:
#           FITIDYYYYMMDDamt#####, where
#                FITID  = string literal
#                YYYY   = year (numeric)
#                MM     = month (numeric)
#                DD     = day (numeric)
#                amt    = dollar amount of the transaction, including a hypen for negative entries (e.g., -24.95)
#                #####  = 5 digit serial number

#   2.  The 5-digit serial number can change each time you connect to the server, 
#       meaning that the same transaction can download with different FITID numbers.  
#       That's not good, since Money requires a unique FITID value for each valid transaction.  
#       Varying serial numbers result in duplicate transactions!

#   3.  We'll replace the 5-digit serial number with one of our own.  
#       The default will be 0 for every transaction,
#       and we'll increment by one for each subsequent transaction that that matches
#       a previous transaction in the file.

_sD_knownvals = []  #global to keep track of Discover FITID values between regex.sub() calls

def _scrubDiscover(ofx):

    scrubPrint("  +Scrubber: Processing Discover statement.")

    ofx_final = ''      #new ofx message
    _sD_knownvals = []  #reset our global set of known vals (just in case)
    
    #regex p captures everything from <FITID> up to the next <tag>, but excludes the next "<".
    #p produces 2 results:  r.group(1) = <FITID> field, r.group(2)=value
    p = re.compile(r'(<FITID>)([^<\s]+)',re.IGNORECASE)

    #call substitution (inline lamda, takes regex result = r as tuple)
    ofx_final = p.sub(lambda r: _scrubDiscover_r1(r), ofx)

    return ofx_final

def _scrubDiscover_r1(r):
    #regex subsitution function for _scrubDiscover()
    global _sD_knownvals

    fieldtag = r.group(1)
    fitid = r.group(2).strip(' ')
    
    #pointer to end of "base" FITID value
    bx = len(fitid) - 5
    fitid_b = fitid[:bx]
    
    #find a unique serial#, from 0 to 9999
    seq = 0   #default
    while seq < 9999:
        fitid = fitid_b + str(seq)
        exists = (fitid in _sD_knownvals)
        if exists:  #already used it... try another
            seq=seq+1
        else:
            break   #unique value... write it out
        
    _sD_knownvals.append(fitid)         #remember the assigned value between calls
    return fieldtag + fitid             #return the new string for regex.sub()

#--------------------------------    
def _scrubTime(ofx):
    #Replace NULL time stamps with noontime (12:00)

    #regex p captures everything from <DT*> up to the next <tag>, but excludes the next "<".
    #p produces 2 results:  group(1) = <DT*> field, group(2)=dateval
    p = re.compile(r'(<DT.+?>)([^<\s]+)',re.IGNORECASE)
    #call date correct function (inline lamda, takes regex result = r tuple)
    
    nullTimeUpdated = False
    ofx_final = p.sub(lambda r: _scrubTime_r1(r), ofx)
    if nullTimeUpdated: scrubPrint("  +Scrubber: Null time values updated.")
    
    return ofx_final

def _scrubTime_r1(r):
    # Replace zero and NULL time fields with a "NOON" timestamp (120000)
    # Force "date" to be the same as the date listed, regardless of time zone by setting time to NOON.
    # Applies when no time is given, and when time == MIDNIGHT (000000)
    fieldtag = r.group(1)
    DT = r.group(2).strip(' ')      #date+time
    
    # Full date/time format example:  20100730000000.000[-4:EDT]
    if DT[8:] == '' or DT[8:14] == '000000':
        #null time given.  Adjust to 120000 value (noon).
        DT = DT[:8] + '120000'
        nullTimeUpdated = True
        
    return fieldtag + DT

#--------------------------------    
def _scrubDTSTART(ofx):
    # <DTSTART> field for an account statement must have a matching <DTEND> field
    # If DTEND is missing, insert <DTEND>="now"
    # The assumption is made that only one statement exists in the OFX file (no multi-statement files!)
    
    ofx_final = ofx
    now = datetime.datetime.now()
    nowstr = now.strftime("%Y%m%d%H%M00")
    
    if ofx.find('<DTSTART>') >= 0 and ofx.find('<DTEND>') < 0:
        #we have a dtstart, but no dtend... fix it.
        scrubPrint("  +Scrubber: Fixing missing <DTEND> field")
        
        #regex p captures everything from <DTSTART> up to the next <tag> or white space into group(1)
        p = re.compile(r'(<DTSTART>[^<\s]+)',re.IGNORECASE)
        if Debug: print "DTSTART: findall()=", p.findall(ofx_final)
        #replace group1 with (group1 + <DTEND> + datetime)
        ofx_final = p.sub(r'\1<DTEND>'+nowstr, ofx_final)
    
    return ofx_final

def _scrubShiftTime(ofx, h):
    #Shift DTASOF time values by (float) h hours
    #Added: 15-Feb-2011, rlc
    
    #regex p captures everything from <DTASOF> up to the next <tag> or white-space.
    #p produces 2 results:  group(1) = <DTASOF> field, group(2)=dateval
    p = re.compile(r'(<DTASOF>)([^<\s]+)',re.IGNORECASE | re.DOTALL)
    
    #call date correct function (inline lamda, takes regex result = r tuple)
    if p.search(ofx): 
        scrubPrint("  +Scrubber: Shifting DTASOF time values " + str(h) + " hours.")
        ofx_final = p.sub(lambda r: _scrubShiftTime_r1(r,h), ofx)    

    return ofx_final

def _scrubShiftTime_r1(r,h):
    #Shift time value by (float) h hours for regex search result r.
    #Added: 15-Feb-2011, rlc
    
    fieldtag = r.group(1)       #date field tag (e.g., <DTASOF>)
    DT = r.group(2).strip(' ')  #date+time

    if Debug: print "fieldtag=", fieldtag, "| DT=" + DT
    
    # Full date/time format example:  20100730120000.000[-4:EDT]
    #separate into date/time + timezone
    tz = ""
    if '[' in DT:
        p = DT.index('[')
        tz = DT[p:]
        DT = DT[:p]
    
    #strip the decimal fraction, if we have it
    if '.' in DT:
        d  = DT.index('.')
        DT = DT[:d]
        
    if Debug: scrubPrint("New DT=" + DT + "| tz=" + tz)
    
    #shift the time
    tval = datetime.datetime.strptime(DT,"%Y%m%d%H%M%S")  #convert str to datetime
    deltaT = datetime.timedelta(hours=h)    
    tval += deltaT                                        #add hours
    DT = tval.strftime("%Y%m%d%H%M%S") + tz               #convert new datetime to str
        
    return fieldtag + DT

def _scrubINVsign(ofx):
    #Fix malformed parameters in Investment buy/sell sections, if they exist
    #Issue  first noticed with Fidelity netbenefits 401k accounts:  rlc*2013
    
    #BUY transactions:
    #   UNITS must be positive
    #   TOTAL must be negative
    
    #SELL transactions:
    #   UNITS must be negative
    #   TOTAL must be positive
    
    stat=False
    p = re.compile(r'(<INVBUY>|<INVSELL>)(.+?<UNITS>)(.+?)(<.+?<TOTAL>)([^<\r\n]+)', re.IGNORECASE)
    ofx_final=p.sub(lambda r: _scrubINVsign_r1(r), ofx)
    if stat:
        scrubPrint("  +Scrubber: Invalid investment sign (pos/neg) found.  Corrected.")
    
    return ofx_final
    
def _scrubINVsign_r1(r):
    type=""
    if "INVBUY" in r.group(1): type = "INVBUY"
    if "INVSELL" in r.group(1): type="INVSELL"
    qty = r.group(3)
    total=r.group(5)
    qty_v=0
    total_v=0
    try:
        qty_v=float(qty)
        total_v=float(total)
    except:
        pass
    
    if (type=="INVBUY" and qty_v<0) or (type=="INVSELL" and qty_v>0):
        stat=True
        qty=str(-1*qty_v)

    if (type=="INVBUY" and total_v>0) or (type=="INVSELL" and total_v<0):
        stat=True
        total=str(-1*total_v)
    
    rtn=r.group(1) + r.group(2) + qty + r.group(4) + total
  
    return rtn

def _scrubGeneral(ofx):    
    # General scrub routine for singular tag substitutions 
    # Remove tag/value pairs that Money doesn't support
    
    #define unsupported tags that we've had trouble with
    uTags = []    
    uTags.append('CORRECTACTION')
    uTags.append('CORRECTFITID')
    
    # remove tag/value pairs from ofx
    for tag in uTags:
        p = re.compile(r'<'+tag+'>[^<]+',re.IGNORECASE) 
        if p.search(ofx):
            ofx = p.sub('',ofx)
            print("  +Scrubber: <"+tag+"> tags removed.  Not supported by Money.")
    
    return ofx
