from datetime import datetime
from datetime import timedelta
from dateutil.relativedelta import relativedelta
from pytz import timezone
from yahooquery import Ticker
import yahooquery
import pandas as pd
from forex_python.converter import CurrencyRates
import forex_python
currencyConverter = CurrencyRates()
import statistics



class StockDataForDay:
    def __init__(self, stockSymbol=None, date=None):
        if stockSymbol is None:
            stockSymbol = 'ERROR'
        self.stockSymbol = stockSymbol

        if date is None:
            #date = datetime.now(timezone('US/Eastern'))
            date = 'ERROR'
        self.desiredDateForStockData = date
        self.actualTickerRegMarketTimeDate = None

        self.name = 'ERROR'

        self.closePrice = float('nan')
        self.previousClosePrice = float('nan')
        self.dayPriceChangeDollars = float('nan')
        self.priceChangePercent = float('nan')
        self.peRatio = float('nan') #unstable if calculating for older dates

        self.quarterDateUpToFourQ = []
        self.netMarginUpToFourQ = []
        self.revenueUpToFourQ = []
        self.netProfitUpToFourQ = []

        self.financialCurrency = ''
        self.conversionFactorToConvertToUSD = []

        self.marketCap = float('nan')
        self.volume = float('nan')
        self.dividendRate = float('nan')
        self.dividendYield = float('nan')

        self.sortingValue = float('nan')

        self.meanPriceForMonthFiveYearsAgo = float('nan')
        self.meanPriceForMonthTenYearsAgo = float('nan')
        self.closePriceMaxAgo = float('nan')

        self.intersectionData = None



def getMeanPriceYearsAgo(yrsAgo,startDate, tickerYahooQueryStock):

    priceExistsNumYearsAgo = True
    dateNumYearsAgo = startDate - relativedelta(years=yrsAgo)
    stockHistoryNumYearsAgo = tickerYahooQueryStock.history(start=dateNumYearsAgo - timedelta(days=15), end=dateNumYearsAgo + timedelta(days=16))

    if isinstance(stockHistoryNumYearsAgo, pd.DataFrame):
        closePricesNumYearsAgo = stockHistoryNumYearsAgo.loc[:, 'close'].to_numpy()
        meanPriceForMonthNumYearsAgo = statistics.mean(closePricesNumYearsAgo)
        return meanPriceForMonthNumYearsAgo, priceExistsNumYearsAgo
    else:
        stockHistoryMaxAgo = tickerYahooQueryStock.history(period='max', interval='1d')
        closePriceMaxAgo = stockHistoryMaxAgo.loc[stockHistoryMaxAgo.index[0], 'close']
        priceExistsNumYearsAgo = False
        return closePriceMaxAgo, priceExistsNumYearsAgo


def getStockData(stockSymbolList, date, listOfAllQuarterDates=[]):#date is kind of irrelevant if I'm limited to analyzing same day summary data

    listOfStockData = []

    i=0
    for stockSymbol in stockSymbolList:
        stock = StockDataForDay(stockSymbol, date)

        print(stock.stockSymbol + ' data retrieval started for date: ' + date.strftime('%m-%d-%Y'))
        '''
        When trying to get data for multiple days, I run into a request limit pretty easily
        Can either: 1) Increase backoff_factor or 2) reduce yahooQueryStock calls i.e. requests.
        Backoff factor of 10 takes 4-5minutes (4.5min) to resume and got all the way through 3 days.
        Backoff factor of 7 takes 3-4minutes (3.5min) to resume and got all the way through 3 days, but not 4 days worth of data.
        '''
        yahooQueryStock = Ticker(stockSymbol, status_forcelist=[404,429,500,502,503,504], backoff_factor=10, timeout=None)#if making a lot of requests, I might get a 404 error, this tells it to retry I think. Some other parameters dealing with many requests seem useful such as backoff_factor. Check keyword arguments section of documention
        stockQuote_typeData = yahooQueryStock.quote_type
        stock.name = stockQuote_typeData[stockSymbol]['shortName']
        stockPriceData = yahooQueryStock.price

        stockPriceDataRegularMarketTimeAsString = stockPriceData[stockSymbol]['regularMarketTime']
        stock.actualTickerRegMarketTimeDate = datetime.strptime(stockPriceDataRegularMarketTimeAsString,"%Y-%m-%d %H:%M:%S")#create datetime object from string

        summaryDetailData = yahooQueryStock.summary_detail
        if ('marketCap' in summaryDetailData[stockSymbol]):
            stock.marketCap = summaryDetailData[stockSymbol]['marketCap']
        stock.volume = summaryDetailData[stockSymbol]['volume']
        if ('dividendRate' in summaryDetailData[stockSymbol]):
            stock.dividendRate = summaryDetailData[stockSymbol]['dividendRate']
            stock.dividendYield = summaryDetailData[stockSymbol]['dividendYield']


        #this is only necessary if not retrieving data on same day, I can just use the summary data. But I don't think I can access past summary data then can I? Problematic if not running program on same day!!
        oneWeekEarlier = date - timedelta(days=7)#can't simply substract a day to find date of last close because stock market is not open every day. Look back a week and hope that's long enough to find another day market was open
        nextDay = date+timedelta(days=1)
        stockHistoryPastWeek = yahooQueryStock.history(start=oneWeekEarlier, end=nextDay)#this isn't safe because if the stocks were closed for more than 1 week,I would miss previous day
        #not quite sure exactly why I need this next bit (maybe time zone differences?) But 2454.TW data includes date+1 day data for some reason. Notably, time is close to midnight.
        #checks to see if last day in past week data is date + 1 instead of the expected date, and corrects the data as necessary (throws out the date + 1 data)
        if (stockHistoryPastWeek.index[-1][1] == nextDay.date()):
            historyPastThreeDays = stockHistoryPastWeek.tail(3)
            historyTodayAndPreviousDayOnly = historyPastThreeDays.head(2)
        else:
            historyTodayAndPreviousDayOnly = stockHistoryPastWeek.tail(2)

        dateOfInterestAccordingToHistoricalData = historyTodayAndPreviousDayOnly.index[-1][1]  # this gives a datetime date, not a datetime object
        # actualTickerRegMarketTimeDate does not equal datetime.now('US/Eastern').date(). It simply equals the regular market time. The timezone of the time is timezone of os at time of starting python program
        # if the yahoo query ticker data (which is only real time as far as I know) is from the date of interest, assume that the historical data (which is sometimes the case I think, not sure how quickly historical updates after close) may not have today's price, so get current price from summary
        # **********TRY running this code just after close to see if historical data is updated.
        if (stock.actualTickerRegMarketTimeDate.date() == stock.desiredDateForStockData.date()):
            stock.closePrice = stockPriceData[stockSymbol]['regularMarketPrice']
            print('-ticker date matches date of interest')
            '''
            #these lines of code don't work because the regularMarketPreviousClose, regularMarketChange, and regularMarketChangePercent data is sometimes missing or wrong. Example Stock: ^Ndx
            stock.previousClosePrice = stockPriceData[stockSymbol]['regularMarketPreviousClose']
            stock.dayPriceChangeDollars = stockPriceData[stockSymbol]['regularMarketChange']
            stock.priceChangePercent = stockPriceData[stockSymbol]['regularMarketChangePercent']
            '''
        else:
            if(dateOfInterestAccordingToHistoricalData == stock.desiredDateForStockData.date()):
                print('-good date of interest in historical data, but ticker date does not match date of interest. date of interest likely does not equal today')
                stock.closePrice = historyTodayAndPreviousDayOnly.close[-1]
            else:
                print('\n*****\n\nERROR: stock price for date of interest is missing. Something went very wrong. Investigate high priority for '+ stock.stockSymbol + '\n\n*****\n')

        lengthOfHistoricalData = len(historyTodayAndPreviousDayOnly.index)
        if((dateOfInterestAccordingToHistoricalData == stock.desiredDateForStockData.date()) and (lengthOfHistoricalData == 2)):
            stock.previousClosePrice = historyTodayAndPreviousDayOnly.close[0]
            print('-historical data is all there, including date of interest and previous close. Everything is good')
        elif (lengthOfHistoricalData == 2):
            stock.previousClosePrice = historyTodayAndPreviousDayOnly.close[1]
            print('-Historical data might not have been updated to include date of interest yet? all historical data present and can get get previous close from it, but date of interest is not in historical data.')
        elif ((dateOfInterestAccordingToHistoricalData != stock.desiredDateForStockData.date()) and lengthOfHistoricalData == 1):
            stock.previousClosePrice = historyTodayAndPreviousDayOnly.close[0]
            print('\n*****\n\nWARNING: some historical data is missing, but there is one data point and it might be previous close. May or may not be ok. Stock: ' + stock.stockSymbol + '\n\n*****\n')
        else:
            print('\n*****\n\nERROR: something is very wrong with historical data, probably missing. Investigate Stock: ' + stock.stockSymbol + '\n\n*****\n')

        stock.dayPriceChangeDollars = stock.closePrice - stock.previousClosePrice
        stock.priceChangePercent = stock.dayPriceChangeDollars / stock.previousClosePrice


        #this may not be accurate if not running on same day
        #stockSummary = yahooQueryStock.summary_detail
        #stock.peRatio = stockSummary[stockSymbol]['trailingPE']
        try:
            stockTTM_EPS = yahooQueryStock.quotes[stockSymbol.upper()]['epsTrailingTwelveMonths']
            stock.peRatio = stock.closePrice / stockTTM_EPS
        except KeyError:
            print('\n*****\n\nKeyError ERROR (I expect this, but want to get rid of it):, epsTrailingTwelveMonths data does not exist. Either index fund or data missing for stock:' + stock.stockSymbol + '\n\n*****\n')
            #pass#index funds dont have PE ratios

        try:
            # this may not be accurate if not running on same day
            netIncome = yahooQueryStock.get_financial_data('NetIncome', 'q', trailing=False)
            totalRevenue = yahooQueryStock.get_financial_data('TotalRevenue', 'q', trailing=False)

            if (isinstance(netIncome,str) and isinstance(totalRevenue, str)):
                print('\n\n' + stock.stockSymbol + ' is not a normal stock? Maybe index fund. Either that or net income and total revenue both missing' + '\n\n')
            elif(isinstance(netIncome,str) or isinstance(totalRevenue, str)):
                print('\n\n' + stock.stockSymbol + ' has partial financial data. One of net income or total revenue is present, but the other is missing.' + '\n\n')
            elif (isinstance(netIncome, dict) or isinstance(totalRevenue,dict)):
                #can maybe fix this by using status_forcelist paramter when creating Ticker only or can I use it elsewhere?
                print('***************\n********************\n*************\n netIncome and/or totalRevenue is dict. Probably some sort of request error like 404. Stock ' + stock.stockSymbol +
                      '\n***************\n********************\n*************\n')
                print('netIncome:')
                print(netIncome)
                print('totalRevenue:')
                print(totalRevenue)
                print('stop here to see what going on')
            else:

                stock.financialCurrency = yahooQueryStock.financial_data[stockSymbol]['financialCurrency']

                #filter data to only get data of period type 3M (quarterly). some data includes 1 or more TTM data for season
                netIncomeFiltered = netIncome[(netIncome['periodType'] == '3M')]
                totalRevenueFiltered = totalRevenue[(totalRevenue['periodType'] == '3M')]

                netIncomeSortedAndFiltered = netIncomeFiltered.sort_values('asOfDate',ascending=False)
                totalRevenueSortedAndFiltered = totalRevenueFiltered.sort_values('asOfDate', ascending=False)

                i = 1#i don't think this is necessary
                for row in netIncomeSortedAndFiltered.loc[:, ['asOfDate','NetIncome']].itertuples(index=False):
                    if (i <= 4):
                        stock.quarterDateUpToFourQ.append(row.asOfDate)
                        if (stock.financialCurrency != 'USD'):
                            try:
                                stock.conversionFactorToConvertToUSD.append(currencyConverter.convert(stock.financialCurrency,'USD',1,row.asOfDate))
                            except forex_python.converter.RatesNotAvailableError:
                                endOfMonthCurrencyDataAllTime = yahooquery.currency_converter(stock.financialCurrency, "USD", "allTime")
                                for iterData in reversed(endOfMonthCurrencyDataAllTime['HistoricalPoints']):
                                    testDate = datetime.strptime(iterData['PointInTime'], '%Y-%m-%d %H:%M:%S')
                                    if(testDate.date() == row.asOfDate.date()):
                                        stock.conversionFactorToConvertToUSD.append( iterData['InterbankRate'] )
                                        break



                        if ((row.asOfDate in listOfAllQuarterDates) == False):
                            listOfAllQuarterDates.append(row.asOfDate)
                        stock.netProfitUpToFourQ.append(row.NetIncome)
                        i = i + 1#i don't think this is necessary

                i = 1
                for row in totalRevenueSortedAndFiltered.loc[:, ['asOfDate', 'TotalRevenue']].itertuples(index=False):
                    if (i <= 4):
                        stock.revenueUpToFourQ.append(row.TotalRevenue)
                        i = i + 1

                for i in range(len(stock.netProfitUpToFourQ)):
                    stock.netMarginUpToFourQ.append( stock.netProfitUpToFourQ[i] / stock.revenueUpToFourQ[i] )


        #except AttributeError:
            #pass#Index funds don't have this data. Do nothing, null variables already exist.
        #should probably do my own error checking to get rid of this. See below in the March History code where I check if returned data is a dataframe
        except:
            print('\n*****\n\nTypeError ERROR: There was a TypeError when getting financial data. Something unexpected went very wrong. Look into this for stock: ' + stock.stockSymbol + '\n\n*****\n')
            #pass#Index funds don't have this data. Do nothing, null variables already exist.



        #Find March Low price and calculate bounce from the March low [(current price - march low)/divided by march low].
        stockHistoryMarch2020 = yahooQueryStock.history(start=datetime(2020,3,1), end=datetime(2020,4,1))
        #check if there is March data (i.e. a DataFrame will be returned).
        #If there isn't March data, a string saying data unavailable is returned, in which case I leave the sorting value as the default nan
        if isinstance(stockHistoryMarch2020, pd.DataFrame):
            closePricesMarch2020 = stockHistoryMarch2020.loc[:, 'close'].to_numpy()
            marchLowClosePrice = min(closePricesMarch2020)
            stock.sortingValue = (stock.closePrice - marchLowClosePrice)/marchLowClosePrice



        price, priceExistsNumYearsAgo = getMeanPriceYearsAgo(5, date, yahooQueryStock)
        if (priceExistsNumYearsAgo == True):
            stock.meanPriceForMonthFiveYearsAgo = price
        else:
            stock.closePriceMaxAgo = price

        price, priceExistsNumYearsAgo = getMeanPriceYearsAgo(10, date, yahooQueryStock)
        if (priceExistsNumYearsAgo == True):
            stock.meanPriceForMonthTenYearsAgo = price
        else:
            stock.closePriceMaxAgo = price







        listOfStockData.append(stock)
        i+=1
        print(stock.stockSymbol + ' data retrieval finished for date: ' + date.strftime('%m-%d-%Y'))


    return(listOfStockData, listOfAllQuarterDates)