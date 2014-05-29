# quotes.py
# http://sites.google.com/site/pocketsense/
#
# Original Version: TFB (http://thefinancebuff.com)
#
# This script retrieves price quotes for a list of stock and mutual fund ticker symbols from Yahoo! Finance.
# It creates a dummy OFX file and then imports the file to the default application associated with the .ofx extension.
# I wrote this script in order to use Microsoft Money after the quote downloading feature is disabled by Microsoft.
#
# For more information, see
#    http://thefinancebuff.com/2009/09/security-quote-script-for-microsoft-money.html

# Revisions (pocketsense)
# -----------------------------------------------------
# 04-Mar-2010*rlc 
#   - Initial changes/edits for incorporation w/ the "pocketsense" pkg (formatting, method of call, etc.)
#   - Examples:
#       - Debug control
#       - use .\xfr for data
#       - moved stock/fund ticker symbols to sites.dat, separating user info from the code
# 18-mar-2010*rlc
#   - Skip stock/fund symbols that aren't recognized by Yahoo! Finance (rather than throw an error)
# 07-May-2010*rlc
#   - Try not to bomb if the server connection fails or times out and set return STATUS accordingly
#   - Changed output file format to QUOTES+time$.ofx
# 09-Sep-2010*rlc
#   - Add support for alternate Yahoo! quote site URL (defined in sites.dat as YahooURL: url)
#   - Use CSV utility to parse csv data
#   - Write out quote history file to quotes.csv
# 12-Oct-2010*rlc
#   - Fixed bug in QuoteHistory date field
# 24-Nov-2010*rlc
#   - Skip quotes with missing parameters.  Otherwise, Money mail throw an error during import.
#   - Add an "account balance" data field of zero (balance=0) for the overall statement.
# 10-Jan-2010*rlc
#   - Write quote summary to xfrdir\quotes.htm
#   - Moved _header, OfxField, OfxTag, _genuid, and OfxDate functios to   Removed "\r" from linebreaks.
#       Will reuse when combining statements
# 18-Feb-2011*rlc:
#   - Use ticker symbol when stock name from Yahoo is null
#   - Added Cal's yahoo screen scraper for symbols not available via the csv interface
#   - Added YahooTimeZone option
#   - Added support for quote multiplier option
# 22-Mar-2011*rlc:
#   - Added support for alternate ticker symbol to send to Money (instead of Yahoo symbol)
#     Default symbol = Yahoo ticker
# 03Sep2013*rlc
#   - Modify to support European versions of MS Money 2005 (and probably 2003/2004)
#     * Added INVTRANLIST tag set
#     * Added support for forceQuotes option.  Force additional quote reponse to adjust 
#       shares held to non-zero and back.
#   - Updated YahooScrape code for updated Yahoo html screen-formatting
# 19Jul2013*rlc
#   - Bug fix.  Strip commas from YahooScrape price results
#   - Added YahooScrape support for quote symbols with a '^' char (e.g., ^DJI)
# 21Oct2013*rlc
#   - Added support for quoteAccount
# 09Jan2014*rlc
#   - Fixed bug related to ForceQuotes and change of calendar year.
# 19Jan2014*rlc:  
#   -Added support for EnableGoogleFinance option
#   -Reworked the way that quotes are retrieved, to improve reliability
# 14Feb2014*rlc:
#   -Extended url timeout.  Some ex-US users were having issues.
#   -Fixed bug that popped up when EnableYahooFinace=No
# 25Feb2014*rlc:
#   -Changed try/catch for URLopen to catch *any* exception

import os, sys, time, urllib2, socket, shlex, re, csv, uuid
import site_cfg
from rlib1 import *
from datetime import datetime
from control2 import *

join = str.join

class Security:
    """
    Encapsulate a stock or mutual fund. A Security has a ticker, a name, a price quote, and 
    the as-of date and time for the price quote. Name, price and as-of date and time are retrieved
    from Yahoo! Finance.
    
    fields:
        status, source, ticker, name, price, quoteTime, pclose, pchange
    """

    def __init__(self, item):
        #item = {"ticker":TickerSym, 'm':multiplier, 's':symbol}
        # TickerSym = symbol to grab from Yahoo
        # m         = multiplier for quote
        # s         = symbol to pass to Money
        self.ticker = item['ticker']
        self.multiplier = item['m']
        self.symbol = item['s']
        self.status = True
        socket.setdefaulttimeout(10)    #default socket timeout for server read, secs
        
    def _removeIllegalChars(self, inputString):
        pattern = re.compile("[^a-zA-Z0-9 ,.-]+")
        return pattern.sub("", inputString)
        
    def getQuote(self):
        
        #Yahoo! Finance:  
        #    name (n), lastprice (l1), date (d1), time(t1), previous close (p), %change (p2)
        
        if Debug: print "Getting quote for:", self.ticker
        url = YahooURL+"/d/quotes.csv?s=%s&f=nl1d1t1pp2" % self.ticker
        
        self.status=False
        self.source='Y'
        #note: each try for a quote sets self.status=true if successful
        if eYahoo:
            try:
                csvtxt = urllib2.urlopen(url).read()
                quote = self.csvparse(csvtxt)
                self.quoteURL = YahooURL + '/q?s=%s&ql=1url' % self.ticker
            except:
                print "** An error occurred when connecting to the Yahoo CSV service"
                self.status = False
        
            if not self.status and eYScrape:
                # try screen scrape
                csvtxt = self.YahooScrape()
                quote = self.csvparse(csvtxt)        
        
        if not self.status and eGoogle:
            # try screen scrape
            csvtxt = self.GoogleScrape()
            quote = self.csvparse(csvtxt)
            if self.status: self.source='G'
            
        if not self.status:
            print "** ", self.ticker, ': invalid quote response. Skipping...'
            self.name = '*InvalidSymbol*'
        else:
            #show/save what we got   rlc*2010
            # example: "Amazon.com, Inc.",78.46,"9/3/2009","4:00pm", 80.00, "-1.96%"
            # Security names may have embedded commas, so use CSV utility to parse (rlc*2010)

            if Debug: print "Quote result string:", csvtxt

            self.name    = quote[0]
            self.price   = quote[1]
            self.date    = quote[2]
            self.time    = quote[3]
            self.pclose  = quote[4]
            self.pchange = quote[5]
            
            #clean things up, format datetime str, and apply multiplier
            #ampersand character (&) is not valid in OFX
            self.name = self._removeIllegalChars(self.name)
            # if security name is null, replace with name with symbol
            if len(self.name.replace(" ", ""))==0: self.name = self.ticker
            self.price = str(float2(self.price)*self.multiplier)  #adjust price by multiplier
            self.date = self.date.lstrip('0 ')
            self.datetime  = datetime.strptime(self.date + " " + self.time, "%m/%d/%Y %I:%M%p")
            self.quoteTime = self.datetime.strftime("%Y%m%d%H%M%S") + '[' + YahooTimeZone + ']'
            if '?' not in self.pclose and 'N/A' not in self.pclose:
                #adjust last close price by multiplier
                self.pclose = str(float2(self.pclose)*self.multiplier)    #previous close

            name = self.ticker
            if self.symbol <> self.ticker:
                name = self.ticker + '(' + self.symbol + ')'
            print self.source+':' , name, self.price, self.date, self.time

                
    def csvparse(self, csvtxt):
        quote=[]
        self.status=True
        csvlst = [csvtxt]      # csv.reader reads lists the same as files
        reader = csv.reader(csvlst)
        for row in reader:     # we only have one row... so read it into a quote list
            quote = row        # quote[]= [name, price, quoteTime, pclose, pchange], all as strings
       
        if len(quote) < 6:
            self.status=False
        elif quote[1] == '0.00' or quote[2] == 'N/A':
            self.status=False

        return quote

    def YahooScrape(self):
        #New screen scrape function: 02Sep2013*rlc
        #   This function creates a csvtxt string with the same format as the Yahoo csv interface
        #   as with all screen scrapers... any change to the html coding could introduce a glitch.
        
        self.status = True  # fresh start
        csvtxt = ""
        if Debug: print "Trying Yahoo scrape for: ", self.ticker

        #http://finance.yahoo.com/q?s=F0CAN05MQI.TO&ql=1
        url = YahooURL+"/q?s=" + self.ticker+"&ql=1"
        try:
            ht=urllib2.urlopen(url).read().upper()
            self.quoteURL = url
        except:
            print "** error reading " + url + "\n"
            self.status = False
        
        ticker = self.ticker.replace("^","\^")  #use literal regex character
        if self.status:
            # example return: "Amazon.com, Inc.","78.46","9/3/2009","4:00pm", 80.00, "-1.96%"
            try:
                #Name
                t1 = '(<DIV CLASS="TITLE"><H2>)(.*?)(<)'
                p = re.compile(t1)
                rslt = p.findall(ht)
                name= rslt[0][1]
            
                #Last price
                t1 = '(<SPAN ID="YFS_L10_' + ticker + '">)(.*?)(</SPAN>)'
                p = re.compile(t1)
                rslt = p.findall(ht)
                price= rslt[0][1]
                
                #Date: Currently supports format as MMM dd  (Aug 30) or dd MMM (30 Aug).
                #If a time is included, then cheat and assume that time is "now", since 
                #Yahoo servers provide date/time combos in local formats
                
                #note that the span id is repeated for the date
                t = '<SPAN ID="YFS_T10_' + ticker + '">'
                t1 = '(' + t + t + ')(.*?)(</SPAN>)'
                p = re.compile(t1)
                rslt = p.findall(ht)
                date1= rslt[0][1]
                
                tnow   = datetime.now()  	
                yr     = tnow.year
                if ':' in date1:
                    qdate = tnow
                else:
                    try:
                        qdate = datetime.strptime(date1, "%b %d")
                    except:
                        qdate = datetime.strptime(date1, "%d %b")  #throws outer exception if fails

                    qdate = datetime(yr, qdate.month, qdate.day)
                
                #Sanity check on the date. Can't be in the future. Adjust to "last year" if needed
                if qdate > tnow: qdate = datetime(yr-1,qdate.month,qdate.day)
                
                #write it out
                date2 = qdate.strftime("%m/%d/%Y").lstrip('0 ')  #mm/dd/yyyy, but no leading zero or spaces
                #note: price may contain commas, but we don't want them
                price = ''.join([c for c in price if c in '1234567890.'])
                csvtxt = '"' + name + '",' + price + ',"' + date2 + '","4:00pm",?,?'
                if Debug: print "Screen scrape csvtxt=",csvtxt
            except:
                self.status=False
                if Debug: print "Error during screen scrape attempt for: ", self.ticker
        return csvtxt

    def GoogleScrape(self):
        #New screen scrape function: 19-Jan-2014*rlc
        #  This function creates a csvtxt string with the same format as the Yahoo csv interface
        #  Example return: "Amazon.com, Inc.","78.46","9/3/2009","4:00pm", 80.00, "-1.96%"
        #  Gets data from Google Finance
        
        self.status = True  # fresh start
        csvtxt = ""
        if Debug: print "Trying Google Finance for: ", self.ticker

        #Example url:  https://www.google.com/finance?q=msft
        url = GoogleURL + "?q=" + self.ticker
        try:
            ht=urllib2.urlopen(url).read().upper()
            self.quoteURL = url
        except:
            print "** error reading " + url + "\n"
            self.status = False
        
        ticker = self.ticker.replace("^","\^")  #use literal regex character
        if self.status:
            try:
                #Name
                t1 = '(<meta itemprop="name".*?content=")(.*?)(")'
                p = re.compile(t1, re.IGNORECASE | re.DOTALL)
                rslt = p.findall(ht)
                name= rslt[0][1]
            
                #Last price
                t1 = '(<meta itemprop="price".*?content=")(.*?)(")'
                p = re.compile(t1, re.IGNORECASE | re.DOTALL)
                rslt = p.findall(ht)
                price= rslt[0][1]

                #Price change%
                t1 = '(<meta itemprop="priceChangePercent".*?content=")(.*?)(")'
                p = re.compile(t1, re.IGNORECASE | re.DOTALL)
                rslt = p.findall(ht)
                pchange= rslt[0][1] + '%'
                
                #Google date/time format= "yyyy-mm-ddThh:mm:ssZ"
                #               Example = "2014-01-10T21:30:00Z"
                
                t1 = '(<meta itemprop="quoteTime".*?content=")(.*?)(")'
                p = re.compile(t1, re.IGNORECASE | re.DOTALL | re.MULTILINE)
                rslt = p.findall(ht)
                date1= rslt[0][1]
                
                qdate = datetime.strptime(date1, "%Y-%m-%dT%H:%M:%SZ")
                date2 = qdate.strftime("%m/%d/%Y").lstrip('0 ')  #mm/dd/yyyy, but no leading zero or spaces

                #note: price may contain commas, but we don't want them
                price = ''.join([c for c in price if c in '1234567890.'])
                csvtxt = '"' + name + '",' + price + ',"' + date2 + '","4:00pm",?,' + pchange
                if Debug: print "Google csvtxt=",csvtxt
            except:
                self.status=False
                if Debug: print "Error during Google Finance attempt for: ", self.ticker
        return csvtxt
        
class OfxWriter:
    """
    Create an OFX file based on a list of stocks and mutual funds.
    """
    
    def __init__(self, currency, account, shares, stockList, mfList):
        self.currency = currency
        self.account = account
        self.shares = shares
        self.stockList = stockList
        self.mfList = mfList
        self.dtasof = self.get_dtasof()

    def get_dtasof(self):
        #15-Feb-2011: Use the latest quote date/time for the statement
        today = datetime.today()
        dtasof   = today.strftime("%Y%m%d")+'120000'    #default to today @ noon     
        lastdate = datetime(1,1,1)                      #but compare actual dates to long, long ago...
        for ticker in self.stockList + self.mfList:
            if ticker.datetime > lastdate and not ticker.datetime > today:
                lastdate = ticker.datetime
                dtasof = ticker.quoteTime  
                
        return dtasof
        
    def _signOn(self):
        """Generate server signon response message"""
    
        return OfxTag("SIGNONMSGSRSV1",
                    OfxTag("SONRS",
                         OfxTag("STATUS",
                             OfxField("CODE", "0"),
                             OfxField("SEVERITY", "INFO"),
                             OfxField("MESSAGE","Successful Sign On")
                         ),
                         OfxField("DTSERVER", OfxDate()),
                         OfxField("LANGUAGE", "ENG"),
                         OfxField("DTPROFUP", "20010918083000"),
                         OfxTag("FI", OfxField("ORG", "PocketSense"))
                     )
               )

    def invPosList(self):
        # create INVPOSLIST section, including all stock and MF symbols
        posstock = []
        for stock in self.stockList:
            posstock.append(self._pos("stock", stock.symbol, stock.price, stock.quoteTime))

        posmf = []
        for mf in self.mfList:
            posmf.append(self._pos("mf", mf.symbol, mf.price, mf.quoteTime))
            
        return OfxTag("INVPOSLIST",          
                    join("", posstock),     #str.join("",StrList) = "str(0)+str(1)+str(2)..."
                    join("", posmf))


    def _pos(self, type, symbol, price, quoteTime):
        return OfxTag("POS" + type.upper(),
                   OfxTag("INVPOS",
                       OfxTag("SECID",
                           OfxField("UNIQUEID", symbol),
                           OfxField("UNIQUEIDTYPE", "TICKER")
                       ),
                       OfxField("HELDINACCT", "CASH"),
                       OfxField("POSTYPE", "LONG"),
                       OfxField("UNITS", str(self.shares)),
                       OfxField("UNITPRICE", price),
                       OfxField("MKTVAL", str(float2(price)*self.shares)),
                       #OfxField("MKTVAL", "0"),     #rlc:08-2013
                       OfxField("DTPRICEASOF", quoteTime)
                   )
               )

    def invStmt(self, acctid):
        #write the INVSTMTRS section
        stmt = OfxTag("INVSTMTRS",
                OfxField("DTASOF", self.dtasof),
                OfxField("CURDEF", self.currency),
                OfxTag("INVACCTFROM",
                    OfxField("BROKERID", "PocketSense"),
                    OfxField("ACCTID",acctid)
                ),
                OfxTag("INVTRANLIST",
                    OfxField("DTSTART", self.dtasof),
                    OfxField("DTEND", self.dtasof),
                ),
                self.invPosList()
               )

        return stmt

    def invServerMsg(self,stmt):
        #wrap stmt in INVSTMTMSGSRSV1 tag set
        s = OfxTag("INVSTMTTRNRS",
                    OfxField("TRNUID",ofxUUID()),
                    OfxTag("STATUS",
                        OfxField("CODE", "0"),
                        OfxField("SEVERITY", "INFO")),
                    OfxField("CLTCOOKIE","4"), 
                    stmt)
        return OfxTag("INVSTMTMSGSRSV1", s)
        
    def _secList(self):
        stockinfo = []
        for stock in self.stockList:
            stockinfo.append(self._info("stock", stock.symbol, stock.name, stock.price))

        mfinfo = []
        for mf in self.mfList:
            mfinfo.append(self._info("mf", mf.symbol, mf.name, mf.price))

        return OfxTag("SECLISTMSGSRSV1",
                   OfxTag("SECLIST",
                        join("", stockinfo),
                        join("", mfinfo)
                   )
               )

    def _info(self, type, symbol, name, price):
        secInfo = OfxTag("SECINFO",
                       OfxTag("SECID",
                           OfxField("UNIQUEID", symbol),
                           OfxField("UNIQUEIDTYPE", "TICKER")
                       ),
                       OfxField("SECNAME", name),
                       OfxField("TICKER", symbol),
                       OfxField("UNITPRICE", price),
                       OfxField("DTASOF", self.dtasof)
                   )
        if type.upper() == "MF":
            info = OfxTag(type.upper() + "INFO",
                       secInfo,
                       OfxField("MFTYPE", "OPENEND")
                   )
        else:
            info = OfxTag(type.upper() + "INFO", secInfo)

        return info
        
    def getOfxMsg(self):
        #create main OFX message block
        return join('', [OfxTag('OFX',
                        '<!--Created by PocketSense scripts for Money-->',
                        '<!--https://sites.google.com/site/pocketsense/home-->',
                        self._signOn(),
                        self.invServerMsg(self.invStmt(self.account)),
                        self._secList()
                    )])

    def writeFile(self, name):
        f = open(name,"w")
        f.write(OfxSGMLHeader())
        f.write(self.getOfxMsg())
        f.close()

#----------------------------------------------------------------------------
def getQuotes():

    global YahooURL, eYahoo, eYScrape, GoogleURL, eGoogle, YahooTimeZone
    status = True    #overall status flag across all operations (true == no errors getting data)
    
    #get site and other user-defined data
    userdat = site_cfg.site_cfg()
    stocks = userdat.stocks
    funds = userdat.funds
    eYahoo = userdat.enableYahooFinance
    YahooURL = userdat.YahooURL
    GoogleURL = userdat.GoogleURL
    eYScrape = userdat.enableYahooScrape
    eGoogle = userdat.enableGoogleFinance
    YahooTimeZone = userdat.YahooTimeZone
    currency = userdat.quotecurrency
    account = userdat.quoteAccount
    ofxFile1, ofxFile2, htmFileName = '','',''
    
    stockList = []
    print "Getting security and fund quotes..."
    for item in stocks:
        sec = Security(item)
        sec.getQuote()
        status = status and sec.status
        if sec.status: stockList.append(sec)
        
    mfList = []
    for item in funds:
        sec = Security(item)
        sec.getQuote()
        status = status and sec.status
        if sec.status: mfList.append(sec)
        
    qList = stockList + mfList
    
    if len(qList) > 0:        #write results only if we have some data
        #create quotes ofx file  
        if not os.path.exists(xfrdir):
            os.mkdir(xfrdir)
        
        ofxFile1 = xfrdir + "quotes" + OfxDate() + str(random.randrange(1e5,1e6)) + ".ofx"
        writer = OfxWriter(currency, account, 0, stockList, mfList)
        writer.writeFile(ofxFile1)

        if userdat.forceQuotes:
           #generate a second file with non-zero shares.  Getdata and Setup use this file
           #to force quote reconciliation in Money, by sending ofxFile2, and then ofxFile1
           ofxFile2 = xfrdir + "quotes" + OfxDate() + str(random.randrange(1e5,1e6)) + ".ofx"
           writer = OfxWriter(currency, account, 0.001, stockList, mfList)
           writer.writeFile(ofxFile2)
        
        if glob.glob(ofxFile1) == []:
            status = False

        # write quotes.htm file
        htmFileName = QuoteHTMwriter(qList)
        
        #append results to QuoteHistory.csv if enabled
        if status and userdat.savequotehistory:
            csvFile = xfrdir+"QuoteHistory.csv"
            print "Appending quote results to {0}...".format(csvFile)
            newfile = (glob.glob(csvFile) == [])
            f = open(csvFile,"a")
            if newfile:
                f.write('Symbol,Name,Price,Date/Time,LastClose,%Change\n')
            for s in qList:
                #Fieldnames: symbol, name, price, quoteTime, pclose, pchange
                t = s.quoteTime
                t2 = t[4:6]+'/'+t[6:8]+'/'+t[0:4]+' '+ t[8:10]+":"+t[10:12]+":"+t[12:14]
                line = '"{0}","{1}",{2},{3},{4},{5}\n' \
                        .format(s.symbol, s.name, s.price, t2, s.pclose, s.pchange)
                f.write(line)
            f.close()
        
    return status, ofxFile1, ofxFile2, htmFileName