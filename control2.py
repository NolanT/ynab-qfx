# control2.py
# http://sites.google.com/site/pocketsense/
# contains some common configuration data and modules for the ofx pkg
# Initial version: rlc: Feb-2010
#
# 27-Jul-2013: rlc
#   - Added locale support 
#------------------------------------------------------------------------------------

#---MODULES---
import os, sys, pyDes, glob, pickle, locale

#04-Jan-2010*rlc
#   - Added DefaultAppID and DefaultAppVer

Debug = False             #debug mode = true only when testing
#Debug = True

AboutTitle    = 'PocketSense OFX Download Python Scripts'
AboutVersion  = '25-Feb-2014'
AboutSource   = 'http://sites.google.com/site/pocketsense'
AboutName     = 'Robert'

xfrdir   = '.\\xfr\\'        #temp directory for statement downloads
cfgFile  = 'ofx_config.cfg'  #user account settings (can be encrypted)

DefaultAppID  = 'QWIN'
DefaultAppVer = '2200'

if Debug:
    import traceback

def get_int(prompt):
    #get number entry
    prompt = prompt.rstrip() + ' '
    done = False
    while not done:
        istr = raw_input(prompt)
        if istr == '':
            a = 0
            done = True
        else:
            try:
                a=int(istr)
                done = True
            except:
                print 'Please enter a valid integer'
    return a

def FieldVal(dic, fieldname):
    #get field value from a dict list
    #return value for fieldname (returns as type defined in dict)
    val = ''
    fieldname = fieldname.upper()
    if fieldname in dic:
        val = dic[fieldname]
    return val

def decrypt_pw(pwkey):
    #validate password if pwkey isn't null
    if pwkey <> '':
        #file encrypted... need password
        pw = pyDes.getDESpw()   #ask for password
        k = pyDes.des(pw)       #create encryption object using key
        pws = k.decrypt(pwkey,' ')  #decrypt
        if pws <> pw:               #comp to saved password
            print 'Invalid password.  Exiting.'
            sys.exit()
        else:
            #decrypt the encrypted fields
            pwkey = pws
    return pwkey
       
def acctEncrypt(AcctArray, pwkey):
    #encrypt accounts
    d = pyDes.des(pwkey)
    for acct in AcctArray:
       acct[1] = d.encrypt(acct[1],' ')
       acct[3] = d.encrypt(acct[3],' ')
       acct[4] = d.encrypt(acct[4],' ')
    return AcctArray
    
def acctDecrypt(AcctArray, pwkey):
    #decrypt accounts
    d = pyDes.des(pwkey)
    for acct in AcctArray:
       acct[1] = d.decrypt(acct[1],' ')
       acct[3] = d.decrypt(acct[3],' ')
       acct[4] = d.decrypt(acct[4],' ')
    return AcctArray
    
def get_cfg():
    #read in user configuration
    
    c_AcctArray = []        #AcctArray = [['SiteName', 'Account#', 'AcctType', 'UserName', 'PassWord'], ...]
    c_pwkey=''              #default = no encryption
    c_getquotes = False     #default = no quotes
    if glob.glob(cfgFile) <> []:
        cfg = open(cfgFile,'rb')
        try:
            c_pwkey = pickle.load(cfg)            #encrypted pw key
            c_getquotes = pickle.load(cfg)        #get stock/fund quotes?
            c_AcctArray = pickle.load(cfg)        #
        except:
            pass    #nothing to do... must not be any data in the file
        cfg.close()
    return c_pwkey, c_getquotes, c_AcctArray
    


