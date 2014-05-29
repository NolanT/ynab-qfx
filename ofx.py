#ofx.py
# http://sites.google.com/site/pocketsense/

# Original version: by Steve Dunham

# Revisions
# ---------
# 2009: TFB @ "TheFinanceBuff.com"

# Feb-2010*rlc (pocketsense) 
#       - Modified use of the code (call methods, etc.). Use getdata.py to call this routine for specific accounts.
#       - Added scrubber module to clean up known issues with statements. (currently only Discover)
#       - Modified sites structure to include minimum download period
#       - Moved site, stock, fund and user-defined parameters to sites.dat and implemented a parser
#       - Perform a bit of validation on output files before sending to Money
#       - Substantial script edits, so that users shouldn't have a need to debug/edit code.

# 07-May-2010*rlc
#   - Try not to bomb if the server connection fails or times out and set return STATUS accordinging

# 10-Sep-2010*rlc
#   - Added timeout to https call

# 12-Oct-2010*rlc
#   - Fixed bug w/ '&' character in SiteName entries

# 30-Nov-2010*rlc
#   - Catch missing <SECLIST> in statements when a <INVPOSLIST> section exists.  This appears to be a required
#     pairing, but sometimes Vanguard to omits the SECLIST when there are no transactions for the period.
#     Money bombs royally when it happens...

# 01-May-2011*rlc
#   - Replaced check for (<INVPOSLIST> & <SECLIST>) pair with a check for (<INVPOS> & <SECLIST>)

# 18Aug2012*rlc
#   - Added support for OFX version 103 and ClientUID parameter.  The version is defined in sites.dat for a 
#     specific site entry, and the ClientUID is auto-generated for the client and saved in sites.dat

# 20Aug2012*rlc
#   - Changed method used for interval selection (default passed by getdata)

# 15Mar2013*rlc
#   - Added sanity check for <SEVERITY>ERROR code in server reply

import time, os, sys, httplib, urllib2, glob, random
import getpass, scrubber, site_cfg
from rlib1 import *
from control2 import *

if Debug:
    import traceback

#define some function pointers
join = str.join
argv = sys.argv

#define some globals
userdat = site_cfg.site_cfg()
                                               
class OFXClient:
    """Encapsulate an ofx client, site is a dict containg siteuration"""
    def __init__(self, site, user, password):
        self.password = password
        self.status = True
        self.user = user
        self.site = site
        self.ofxver = FieldVal(site,"ofxver")
        self.cookie = 3
        site["USER"] = user
        site["PASSWORD"] = password

    def _cookie(self):
        self.cookie += 1
        return str(self.cookie)

    """Generate signon message"""
    def _signOn(self):
        site = self.site

        clientuid=""
        if "103" in self.ofxver: 
            #include clientuid field only if version=103, otherwise the server may reject the request
            clientuid = OfxField("CLIENTUID",userdat.clientuid)
        
        fidata = [OfxField("ORG",FieldVal(site,"fiorg"))]
        fidata += [OfxField("FID",FieldVal(site,"fid"))]
        return OfxTag("SIGNONMSGSRQV1",
                    OfxTag("SONRQ",
                    OfxField("DTCLIENT",OfxDate()),
                    OfxField("USERID",FieldVal(site,"USER")),
                    OfxField("USERPASS",FieldVal(site,"PASSWORD")),
                    OfxField("LANGUAGE","ENG"),
                    OfxTag("FI", *fidata),
                    OfxField("APPID",FieldVal(site,"APPID")),
                    OfxField("APPVER",FieldVal(site,"APPVER")),
                    clientuid
                    ))

    def _acctreq(self, dtstart):
        req = OfxTag("ACCTINFORQ",OfxField("DTACCTUP",dtstart))
        return self._message("SIGNUP","ACCTINFO",req)

    def _bareq(self, bankid, acctid, dtstart, acct_type):
        site=self.site
        req = OfxTag("STMTRQ",
                OfxTag("BANKACCTFROM",
                OfxField("BANKID",bankid),
                OfxField("ACCTID",acctid),
                OfxField("ACCTTYPE",acct_type)),
                OfxTag("INCTRAN",
                OfxField("DTSTART",dtstart),
                OfxField("INCLUDE","Y"))
                )
        return self._message("BANK","STMT",req)
    
    def _ccreq(self, acctid, dtstart):
        site=self.site
        req = OfxTag("CCSTMTRQ",
              OfxTag("CCACCTFROM",OfxField("ACCTID",acctid)),
              OfxTag("INCTRAN",
              OfxField("DTSTART",dtstart),
              OfxField("INCLUDE","Y")))
        return self._message("CREDITCARD","CCSTMT",req)

    def _invstreq(self, brokerid, acctid, dtstart):
        dtnow = time.strftime("%Y%m%d%H%M%S",time.localtime())
        req = OfxTag("INVSTMTRQ",
              OfxTag("INVACCTFROM",
              OfxField("BROKERID", brokerid),
              OfxField("ACCTID",acctid)),
              OfxTag("INCTRAN",
              OfxField("DTSTART",dtstart),
              OfxField("INCLUDE","Y")),
              OfxField("INCOO","Y"),
              OfxTag("INCPOS",
              OfxField("DTASOF", dtnow),
              OfxField("INCLUDE","Y")),
              OfxField("INCBAL","Y"))
        return self._message("INVSTMT","INVSTMT",req)

    def _message(self,msgType,trnType,request):
        site = self.site
        return OfxTag(msgType+"MSGSRQV1",
               OfxTag(trnType+"TRNRQ",
               OfxField("TRNUID",ofxUUID()),
               OfxField("CLTCOOKIE",self._cookie()),
               request))
    
    def _header(self):
        site = self.site
        return join("\r\n",[ "OFXHEADER:100",
                           "DATA:OFXSGML",
                           "VERSION:" + self.ofxver,
                           "SECURITY:NONE",
                           "ENCODING:USASCII",
                           "CHARSET:1252",
                           "COMPRESSION:NONE",
                           "OLDFILEUID:NONE",
                           "NEWFILEUID:" + ofxUUID(),
                           ""])

    def baQuery(self, bankid, acctid, dtstart, acct_type):
        """Bank account statement request"""
        return join("\r\n",
                    [self._header(),
                     OfxTag("OFX",
                          self._signOn(),
                          self._bareq(bankid, acctid, dtstart, acct_type)
                          )
                    ]
                )
                        
    def ccQuery(self, acctid, dtstart):
        """CC Statement request"""
        return join("\r\n",[self._header(),
                    OfxTag("OFX",
                    self._signOn(),
                    self._ccreq(acctid, dtstart))])

    def acctQuery(self,dtstart):
        return join("\r\n",[self._header(),
                    OfxTag("OFX",
                    self._signOn(),
                    self._acctreq(dtstart))])

    def invstQuery(self, brokerid, acctid, dtstart):
        return join("\r\n",[self._header(),
                    OfxTag("OFX",
                    self._signOn(),
                    self._invstreq(brokerid, acctid, dtstart))])

    def doQuery(self,query,name):
        # urllib doesn't honor user Content-type, use urllib2
        garbage, path = urllib2.splittype(FieldVal(self.site,"url"))
        host, selector = urllib2.splithost(path)
        response=False
        try:
            errmsg= "** An ERROR occurred attempting HTTPS connection to"
            h = httplib.HTTPSConnection(host, timeout=5)

            errmsg= "** An ERROR occurred sending POST request to"
            p = h.request('POST', selector, query, 
                     {"Content-type": "application/x-ofx",
                      "Accept": "*/*, application/x-ofx"}
                     )

            errmsg= "** An ERROR occurred retrieving POST response from"
            #allow up to 30 secs for the server response (it has to assemble the statement)
            h.sock.settimeout(30)      
            response = h.getresponse().read()
            f = file(name,"w")
            f.write(response)
            f.close()
        except Exception as inst:
            self.status = False
            print errmsg, host
            print "   Exception type:", type(inst)
            print "   Exception Val :", inst
            if response:
                print "   HTTPS ResponseCode  :", response.status
                print "   HTTPS ResponseReason:", response.reason

        if h: h.close()
            
#------------------------------------------------------------------------------

def getOFX(account, interval):

    sitename  = account[0]
    acct_num  = account[1]
    acct_type = account[2]
    user      = account[3]
    password  = account[4]

    #get site and other user-defined data
    site = userdat.sites[sitename]
    
    #set the interval (days)
    minInterval = FieldVal(site,'mininterval')    #minimum interval (days) defined for this site (optional)
    if minInterval:
         interval = max(minInterval, interval)    #use the longer of the two
    
    #set the start date/time
    dtstart = time.strftime("%Y%m%d",time.localtime(time.time()-interval*86400))
    dtnow = time.strftime("%Y%m%d%H%M%S",time.localtime())
  
    client = OFXClient(site, user, password)
    print sitename,':',acct_num,": Getting records since: ",dtstart
    
    status = True
    #we'll place ofx data transfers in xfrdir (defined in control2.py).  
    #check to see if we have this directory.  if not, create it
    if not os.path.exists(xfrdir):
        try:
            os.mkdir(xfrdir)
        except:
            print '** Error.  Could not create', xfrdir
            system.exit()
    
    #remove illegal WinFile characters from the file name (in case someone included them in the sitename)
    #Also, the os.system() call doesn't allow the '&' char, so we'll replace it too
    sitename = ''.join(a for a in sitename if a not in ' &\/:*?"<>|()')  #first char is a space
    
    ofxFileSuffix = str(random.randrange(1e5,1e6)) + ".ofx"
    ofxFileName = xfrdir + sitename + dtnow + ofxFileSuffix
    
    try:
        if acct_num == '':
            query = client.acctQuery("19700101000000")       #19700101000000 is just a default DTSTART date/time string
        else:
            caps = FieldVal(site, "CAPS")
            if "CCSTMT" in caps:
                query = client.ccQuery(acct_num, dtstart)
            elif "INVSTMT" in caps:
                #if we have a brokerid, use it.  Otherwise, try the fiorg value.
                orgID = FieldVal(site, 'BROKERID')
                if orgID == '': orgID = FieldVal(site, 'FIORG')
                if orgID == '':
                    msg = '** Error: Site', sitename, 'does not have a (REQUIRED) BrokerID or FIORG value defined.'
                    raise Exception(msg)
                query = client.invstQuery(orgID, acct_num, dtstart)

            elif "BASTMT" in caps:
                bankid = FieldVal(site, "BANKID")
                if bankid == '':
                    msg='** Error: Site', sitename, 'does not have a (REQUIRED) BANKID value defined.'
                    raise Exception(msg)
                query = client.baQuery(bankid, acct_num, dtstart, acct_type)

        SendRequest = True
        if Debug: 
            print query
            print
            ask = raw_input('DEBUG:  Send request to bank server (y/n)?').upper()
            if ask=='N': return False, ''
        
        #do the deed
        client.doQuery(query, ofxFileName)
        if not client.status: return False, ''
        
        #check the ofx file and make sure it looks valid (contains header and <ofx>...</ofx> blocks)
        if glob.glob(ofxFileName) == []:
            status = False  #no ofx file?
        else: 
            f = open(ofxFileName,'r')
            content = f.read().upper()
            f.close
            content = ''.join(a for a in content if a not in '\r\n ')  #strip newlines & spaces
            
            if content.find('OFXHEADER:') < 0 and content.find('<OFX>') < 0 and content.find('</OFX>') < 0:
                #throw exception and exit
                raise Exception("Invalid OFX statement.")
                
            #look for <SEVERITY>ERROR code... rlc*2013
            if content.find('<SEVERITY>ERROR') > 0:
                #throw exception and exit
                raise Exception("OFX message contains ERROR condition")

            #attempted debug of a Vanguard issue... rlc*2010
            #if content.find('<INVPOSLIST>') > -1 and content.find('<SECLIST>') < 0:    #DEBUG: rlc*5/2011
            if content.find('<INVPOS>') > -1 and content.find('<SECLIST>') < 0:
                #An investment statement must contain a <SECLIST> section when a <INVPOSLIST> section exists
                #Some Vanguard statements have been missing this when there are no transactions, causing Money to crash
                #It may be necessary to match every investment position with a security entry, but we'll try to just
                #verify the existence of these section pairs. rlc*9/2010
                raise Exception("OFX statement is missing required <SECLIST> section.")
                
            #cleanup the file if needed
            scrubber.scrub(ofxFileName, site)
        
    except Exception as inst:
        status = False
        print inst
        if glob.glob(ofxFileName) <> []:
           print '**  Review', ofxFileName, 'for possible clues...'
        if Debug:
            traceback.print_exc()
        
    return status, ofxFileName
