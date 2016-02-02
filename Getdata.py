# GetData.py
# http://sites.google.com/site/pocketsense/
# retrieve statements, stock and fund data
# Intial version: rlc: Feb-2010

# Revisions
# ---------
# 11-Mar-2010*rlc
#   - Added "interactive" mode
#   - Download all statements and quotes before beginning upload to Money
#   - Allow stock quotes to be sent to Money before statements (option defined in sites.dat) 

# 09-May-2010*rlc
#   - Download files in the order that they will be sent to Money so that file timestamps are in the same order
#   - Send data to Money using the os.system() call rather than os.startfile(), as this seems 
#     to help force the order when sending files to Money (FIFO)
#   - Added logic to catch failed connections and server timeouts
#   - Added "About" title and version to start

# 05-Sep-2010*rlc
#   - Updated to support spaces in SiteName values in sites.dat
#   - Don't auto-close command window if any error is detected during download operations

# 04-Jan-2011*rlc
#   - Display quotes.htm after download if "ShowQuoteHTM: Yes" defined in sites.dat
#   - Ask to display quotes.htm after download if "ShowQuoteHTM: Yes" defined in sites.dat (overrides ShowQuoteHTM)

# 18-Jan-2011*rlc
#   - Added 0.5 s delay between "file starts", which sends an OFX file to Money

# 23Aug2012*rlc
#   - Added user option to change default download interval at runtime
#   - Added support for combineOFX

# 28Aug2013*rlc
#   - Added support for forceQuotes option

# 21Oct2013*rlc
#   - Modified forceQuote option to prompt for statement accept in Money before continuing

# 25Feb2014*rlc
#   - Bug fix for forceQuote option when the quote feature isn't being used

import os, sys, glob, time
import ofx, quotes, site_cfg
from control2 import *
from rlib1 import *

if __name__=="__main__":

    stat1 = True    #overall status flag across all operations (true == no errors getting data)
    print AboutTitle + ", Ver: " + AboutVersion + "\n"
    
    if Debug: print "***Running in DEBUG mode.  See Control2.py to disable***\n"
    doit = raw_input("Download transactions? (Y/N/I=Interactive) [Y] ").upper()
    if len(doit) > 1: doit = doit[:1]    #keep first letter
    if doit == '': doit = 'Y'
    if doit in "YI":

        userdat = site_cfg.site_cfg()
        
        #get download interval, if promptInterval=Yes in sites.dat
        interval = userdat.defaultInterval
        if userdat.promptInterval:
            try:
                p = int2(raw_input("Download interval (days) [" + str(interval) + "]: "))
                if p>0: interval = p
            except:
                print "Invalid entry. Using defaultInterval=" + str(interval)
        
        #get account info
        #AcctArray = [['SiteName', 'Account#', 'AcctType', 'UserName', 'PassWord'], ...]
        pwkey, getquotes, AcctArray = get_cfg()
        ofxList = []
        quoteFile1, quoteFile2, htmFileName = '','',''

        if len(AcctArray) > 0 and pwkey <> '':
            #if accounts are encrypted... decrypt them
            pwkey=decrypt_pw(pwkey)
            AcctArray = acctDecrypt(AcctArray, pwkey)
    
        #delete old data files
        ofxfiles = xfrdir+'*.ofx'
        if glob.glob(ofxfiles) <> []:
            os.system("del "+ofxfiles)
            
        print "Download interval= {0} days".format(interval)
        
        #create process Queue in the right order
        Queue = ['Accts']
        if userdat.savetickersfirst:
            Queue.insert(0,'Quotes')
        else:
            Queue.append('Quotes')

        for QEntry in Queue:

            if QEntry == 'Accts':
               if len(AcctArray) == 0:
                  print "No accounts have been configured. Run SETUP.PY to add accounts"

               #process accounts
               for acct in AcctArray:
                  status, ofxFile = ofx.getOFX(acct, interval)
                  #status == False if ofxFile doesn't exist
                  stat1 = stat1 and status
                  if status: 
                     ofxList.append([acct[0], acct[1], ofxFile])
                  print ""
                        
            #get stock/fund quotes
            if QEntry == 'Quotes' and getquotes:
                status, quoteFile1, quoteFile2, htmFileName = quotes.getQuotes()
                z = ['Stock/Fund Quotes','',quoteFile1]
                stat1 = stat1 and status
                if glob.glob(quoteFile1) <> []: 
                    ofxList.append(z)
                print ""

                # display the HTML file after download if requested to always do so
                if status and userdat.showquotehtm: os.startfile(htmFileName)

        if len(ofxList) > 0:
            print '\nFinished downloading data\n'
            verify = False
            gogo = 'Y'
            if userdat.combineofx and gogo <> 'V':
                cfile=combineOfx(ofxList)       #create combined file

            if doit == 'I' or Debug:
                gogo = raw_input('Upload online data to Money? (Y/N/V=Verify) [Y] ').upper()
                if len(gogo) > 1: gogo = gogo[:1]    #keep first letter
                if gogo == '': gogo = 'Y'

            if gogo in 'YV':
                if glob.glob(quoteFile2) <> []: 
                    if Debug: print "Importing ForceQuotes statement: " + quoteFile2
                    runFile(quoteFile2)  #force transactions for MoneyUK
                    raw_input('ForceQuote statement loaded.  Accept in Money and press <Enter> to continue.')

                print '\nSending statement(s) to Money...'
                if userdat.combineofx and cfile and gogo <> 'V':
                    runFile(cfile)
                else:
                    for file in ofxList:
                        upload = True
                        if gogo == 'V':
                            #file[0] = site, file[1] = accnt#, file[2] = ofxFile
                            upload = (raw_input('Upload ' + file[0] + ' : ' + file[1] + ' (Y/N) ').upper() == 'Y')
                            
                        if upload: 
                           if Debug: print "Importing " + file[2]
                           runFile(file[2])
                        
                        time.sleep(0.5)   #slight delay, to force load order in Money

            #ask to show quotes.htm if defined in sites.dat
            if userdat.askquotehtm:
                ask = raw_input('Open <Quotes.htm> in the default browser (y/n)?').upper()
                if ask=='Y': os.startfile(htmFileName)  #don't wait for browser close
                    
        else:
            if len(AcctArray)>0 or (getquotes and len(userdat.stocks)>0):
                print "\nNo files were downloaded. Verify network connection and try again later."
            raw_input("Press <Enter> to continue...")
        
        if Debug:
            raw_input("DEBUG END:  Press <Enter> to continue...")
        elif not stat1:
            print "\nOne or more accounts (or quotes) may not have downloaded correctly."
            raw_input("Review and press <Enter> to continue...")
