from datetime import datetime
from analyzeAndOutputStockData import analyzeAndOutputStockData
from emailStockData import emailStockData

import os, time, sys #imports necessary for changing os time and restarting program

def changeSystemTimeToEasternStandardTime():
    '''
    os.environ['TZ'] = 'US/Eastern'
    time.tzset() #this only works for unix operating systems
    '''

    currentSystemTimezone = time.tzname[0]
    if(currentSystemTimezone != 'Eastern Standard Time'):
        os.system("tzutil /s \"Eastern Standard Time\"")#for windows only, changes actual os time!!!! Set it back at the end of program. This accesses cmd shell I think
        #os.execv(__file__, sys.argv) #used to restart in some other case. if one of these os.execv lines fails, try the other
        os.execv(sys.executable, ['python'] + sys.argv)#restart this python file


def changeSystemTimeToPacificStandardTime():
    '''
    os.environ['TZ'] = 'US/Eastern'
    time.tzset() #this only works for unix operating systems
    '''

    currentSystemTimezone = time.tzname[0]
    if (currentSystemTimezone != 'Pacific Standard Time'):
        os.system("tzutil /s \"Pacific Standard Time\"")  # for windows only, changes actual os time!!!! Set it back at the end of program. This accesses cmd shell I think



changeSystemTimeToEasternStandardTime()#this is important if you want the data from yahooquery to be in Eastern Standard Time instead of the timezone of the os this script is running on
#print(time.strftime('%I:%M%p   %m-%d-%Y'))




willEmailStockData = False
#emailRecipients = ['mikef522@gmail.com']
#emailRecipients = ['mikef522@gmail.com', 'internet.m@hotmail.com']

#stockSymbolList = ['tsla','fb','aapl','on','Psi','enph']
#stockSymbolList = ['Li','Byddf','gsl-pb','rds-a','tsla','xpev','Psi']
#stockSymbolList = ['TSLA','tsm','main','On']
#stockSymbolList = ['AMC','GME','BB','BBBY','LGND','Fizz','MAC','SKT','AMCX','TR','PLCE','IRBT']
stockSymbolList = ['intc','asml','tsm']
#stockSymbolList = ['Et','Epd','Enb','Pba','gsl-pb','Knop','Xom','Cvx','Tot','Bp','rds-a']
#stockSymbolList = ['BRK-B', 'JPM', 'BAC', 'WFC', 'C', 'MS', 'BLK', 'GS', 'SCHW', 'AXP', 'SPGI', 'CB', 'TFC', 'CME', 'PNC', 'ICE', 'USB', 'MMC', 'PGR', 'AON', 'COF', 'MCO', 'MET', 'TRV', 'TROW', 'AIG', 'BK', 'ALL', 'MSCI', 'PRU', 'AFL', 'FRC', 'DFS', 'WLTW', 'STT', 'SIVB', 'AMP', 'AJG', 'FITB', 'SYF', 'NTRS', 'MKTX', 'HIG', 'MTB', 'KEY', 'RF', 'CFG', 'NDAQ', 'HBAN', 'PFG', 'CINF', 'RJF', 'L', 'CBOE', 'WRB', 'RE', 'GL', 'LNC', 'CMA', 'IVZ', 'AIZ', 'ZION', 'BEN', 'PBCT', 'UNM']

'''
stockSymbolList = ['TSLA','Enph','Qcln','On','ArkG','Cree','Brks','Lscc','Logi','ArkQ','Pypl','Nvda','Iphi','Mrvl','Avgo','Ter','Low','Mpwr','Nxpi','Lrcx','Asml','Qcom','Mchp','Hei','Amd','Aapl','Tsm',
                   'Entg','Cdns','Qrvo','Klac','Amat','Mu','Snps','Mksi','Swks','Oled','^Sox','Soxl','Xlnx','Ttwo','Fb','Amzn','HD','Slab','QQQ','Adi','^Ndx','Atvi','Nflx','Txn','Crus','Goog','Spy',
                   'Ma','Msft','Ea','Dow','V','Cci','Intc','Nio','Kndi','Byddf','Xpev','Li','Jd','Baba','Tcehy','Bidu','Et','Epd','Enb','Pba','gsl-pb','Knop','Xom','Cvx','Tot','Bp','rds-a','Main',
                   'O','Wpc','Pru','PRG','Regn','Pfe','Mrna','Bntx','Soxx','Smh','Xsd','Psi','Ftxl','Xlk']
'''

#stockSymbolList = ['^Ndx','QQQ','^DJI','SPY','^Sox','Soxl']

'''
stockSymbolList = ['AAPL','ACN','ADBE','ADI','ADP','ADSK','AKAM','AMAT','APH','CRM','CSCO','CTSH','CTXS','FFIV','FIS','FISV','FLIR','GLW','HPQ','IBM','INTC',
                           'INTU','JNPR','KLAC','LRCX','MA','MCHP','MSFT','MSI','MU','NTAP','NVDA','ORCL','PAYX','QCOM','STX','TEL','TXN','V','VRSN','WDC','WU','XLNX','XRX',
                           'AVGO','SWKS','QRVO','PYPL','HPE','GPN','SNPS','AMD','DXC','IT','ANSS','CDNS','IPGP','BR','FLT','ANET','FTNT','KEYS','JKHY','MXIM','LDOS','CDW','NLOK',
                           'NOW','ZBRA','PAYC','TYL','TER','VNT','ENPH','TRMB']
'''

datesOfInterest = [datetime(2021,1,27)]
#datesOfInterest = [datetime(2021,1,19), datetime(2021,1,27)]
#datesOfInterest = [datetime(2021,1,19),datetime(2021,1,20),datetime(2021,1,21),datetime(2021,1,22)]




dailyStockDataExcelFileNames = []
dailyGraphOfStocksPdfFilenames = []
filesToEmail =[]

for date in datesOfInterest:
    stockDataExcelFileName, graphOfStocksPdfFileName = analyzeAndOutputStockData(stockSymbolList,date)
    dailyStockDataExcelFileNames.append(stockDataExcelFileName)
    dailyGraphOfStocksPdfFilenames.append(graphOfStocksPdfFileName)

filesToEmail.extend(dailyStockDataExcelFileNames)
filesToEmail.extend(dailyGraphOfStocksPdfFilenames)

if willEmailStockData == True:
    emailStockData(filesToEmail, emailRecipients, datesOfInterest)




changeSystemTimeToPacificStandardTime()
