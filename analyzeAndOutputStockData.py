import datetime
import yfinance as yf

import pandas_datareader as pdr

import pandas as pd
import smtplib

import schedule
import time

from datetime import datetime
from datetime import timedelta
from pytz import timezone

import math


from stockDataRetrieval import getStockData

import xlwings
import numpy

from graphingStockCode import graphStockDataAndReturnStockDataUpdatedWithIntersectionData


def getQuarterandYearNameFromQuarterEndDate(d):
    if (d < datetime(d.year,3,31)):#years are the same, so I'm only comparing month and day
        return 'Q4 ' + str(d.year-1)
    elif (d < datetime(d.year,6,30)):
        return 'Q1 ' + str(d.year)
    elif (d < datetime(d.year,9,30)):
        return 'Q2 ' + str(d.year)
    elif (d < datetime(d.year,12,31)):
        return 'Q3 ' + str(d.year)
    return 'Q4 ' + str(d.year)






def analyzeAndOutputStockData(stockSymbolList,dateOfInterest):
    getStockDataOutput = getStockData(stockSymbolList, dateOfInterest)

    listOfStockData = getStockDataOutput[0]

    '''
    ahistory = yahooQueryStock.history(start=date - timedelta(days=365 * 2), end=date + timedelta(days=1))
    graphStockData(ahistory, stock.stockSymbol, stock.desiredDateForStockData)
    '''

    graphPdfFileName, listOfStockData = graphStockDataAndReturnStockDataUpdatedWithIntersectionData(listOfStockData, listOfStockData[0].desiredDateForStockData)#default graph period is 1 year
    #print(listOfStockData[0].intersectionData)


    #sorting data and putting nan sorting value data on bottom. could sort the nan data by something else if I really wanted.
    listOfStockDataWithNanSortingValue = [ x for x in listOfStockData if (math.isnan(x.sortingValue) == True) ]
    listOfStockDataWithValidSortingValue = [ x for x in listOfStockData if (math.isnan(x.sortingValue) == False) ]
    listOfStockDataWithValidSortingValue.sort(key=lambda x: x.sortingValue, reverse=True)
    listOfStockDataWithValidSortingValue.extend(listOfStockDataWithNanSortingValue)
    completelySortedListOfStockData = listOfStockDataWithValidSortingValue

    listOfAllQuarterDates = getStockDataOutput[1]
    listOfAllQuarterDatesSorted = sorted(listOfAllQuarterDates, reverse=True)









    #create pandas data frame for all the stock data
    stockData2DListForDataFrame = []

    ##
    ##COLUMN HEADERS
    initialColumnTitles = ['March low change(%)', 'Stock', 'Price', 'Change($)', 'Change(%)', 'PE Ratio']
    stockDataColumnTitlesForDataFrame = initialColumnTitles.copy()

    sortedListOfAllQuarterNames = []

    #listOfAllQuarterDatesSorted = (datetime(2020,12,31),datetime(2020,9,30),datetime(2020,6,30),datetime(2020,3,31),datetime(2019,12,31)) #for testing

    #sort all of the quarterly report dates into a list of appropriate quarter and year names. This list will be sorted since the list of quarter dates is sorted
    for i in range(len(listOfAllQuarterDatesSorted)):
        quarterAndYearName = getQuarterandYearNameFromQuarterEndDate(listOfAllQuarterDatesSorted[i])
        if ((quarterAndYearName in sortedListOfAllQuarterNames) == False):
            sortedListOfAllQuarterNames.append(quarterAndYearName)

    quarterlyDataColumnTitles = []

    for i in range(len(sortedListOfAllQuarterNames)):
        quarterlyDataColumnTitles.append(sortedListOfAllQuarterNames[i] + ' Revenue')
        quarterlyDataColumnTitles.append(sortedListOfAllQuarterNames[i] + ' Profits')
        quarterlyDataColumnTitles.append(sortedListOfAllQuarterNames[i] + ' Margins')

    stockDataColumnTitlesForDataFrame.extend(quarterlyDataColumnTitles)

    currentTimeAndDate = datetime.now(timezone('US/Eastern'))
    additionalColumnTitles = ['Stock', 'Financial Data Currency', 'Conversion factors that were used to convert financial data to USD (left to right, latest to oldest quarter)',
                              currentTimeAndDate.strftime('Latest Market cap (as of %m-%d-%Y)'), currentTimeAndDate.strftime('Latest Volume (as of %m-%d-%Y)'),#these 2 must stay together in this order
                              currentTimeAndDate.strftime('Latest Dividend (as of %m-%d-%Y)'), currentTimeAndDate.strftime('Latest Dividend Yield (as of %m-%d-%Y)'),
                              'last date 50MA up over 200MA', 'last date 50MA down under 200MA',
                              'last date 14MA up over 200MA', 'last date 14MA down under 200MA',
                              'last date price up over 200MA', 'last date price down under 200MA',
                              ]#these 2 must stay together in this order
    stockDataColumnTitlesForDataFrame.extend(additionalColumnTitles)


    ##
    ##ORGANIZE STOCK DATA ROW BY ROW
    for stock in completelySortedListOfStockData:
        rFac = 1000#reduction factor for reducing large numbers like revenue and income

        stockDataRowAsList = []

        ##
        ##ADDING PRE-QUARTERLY DATA
        if (math.isnan(stock.sortingValue) == True):#if no data from march, write N/A
            stockDataRowAsList.append('N/A')
        else:
            stockDataRowAsList.append(stock.sortingValue*100)

        stockDataRowAsList.extend([stock.stockSymbol, stock.closePrice, stock.dayPriceChangeDollars, stock.priceChangePercent*100])

        if (math.isnan(stock.peRatio) == True):#if no pe ratio, write N/A
            stockDataRowAsList.append('N/A')
        else:
            stockDataRowAsList.append(stock.peRatio)

        ##
        ##ADDING QUARTERLY DATA
        quarterlyData = []
        '''
        for i in range(len(listOfAllQuarterDatesSorted)):
            quarterlyData.append('No Data')

        matchingDateIndexesSorted = [key for key, val in enumerate(listOfAllQuarterDatesSorted) if val in set(stock.quarterDateUpToFourQ)]
        for i in range(len(matchingDateIndexesSorted)):
            quarterlyData[matchingDateIndexesSorted[i]] = ['${:<15,.0f}'.format(stock.revenueUpToFourQ[i]/rFac), '${:<15,.0f}'.format(stock.netProfitUpToFourQ[i]/rFac), '{:<+7.2f}%'.format(stock.netMarginUpToFourQ[i]*100)]
        '''
        #initialize quarterly data in row to 'No Data'
        for i in range(len(quarterlyDataColumnTitles)):
            quarterlyData.append('No Data')


        for i in range(len(sortedListOfAllQuarterNames)):
            #check to see if current quarter name exists in data for this stock. if it does, fill in quarterly data for that quarter
            for j in range(len(stock.quarterDateUpToFourQ)):
                if (getQuarterandYearNameFromQuarterEndDate(stock.quarterDateUpToFourQ[j]) == sortedListOfAllQuarterNames[i]):
                    #length of sortedListOfAllQuarterNames is 1/3 that of quarterlyDataColumnTitles, so account for that difference in filling in revenue, profit, and margin data
                    #perform currency conversion as necessary
                    if ((stock.financialCurrency != 'USD') and (stock.financialCurrency != '') and (len(stock.conversionFactorToConvertToUSD) != 0)):
                            quarterlyData[i*3] = (stock.revenueUpToFourQ[j]*stock.conversionFactorToConvertToUSD[j]) / rFac
                            quarterlyData[(i*3)+1] = (stock.netProfitUpToFourQ[j]*stock.conversionFactorToConvertToUSD[j]) / rFac

                    else:
                        quarterlyData[i * 3] = stock.revenueUpToFourQ[j] / rFac
                        quarterlyData[(i * 3) + 1] = stock.netProfitUpToFourQ[j] / rFac

                    quarterlyData[(i * 3) + 2] = stock.netMarginUpToFourQ[j] * 100


        stockDataRowAsList.extend(quarterlyData)



        ##
        ##ADDING ADDITIONAL DATA
        stockDataRowAsList.append(stock.stockSymbol)
        if (stock.financialCurrency == ''):
            stockDataRowAsList.append('N/A')
            stockDataRowAsList.append('N/A')
        else:
            if (stock.financialCurrency == 'USD'):
                stockDataRowAsList.append(stock.financialCurrency)
            elif (len(stock.conversionFactorToConvertToUSD) == 0):
                stockDataRowAsList.append(stock.financialCurrency + ' NO CONVERSION AVAILABLE, NOT CONVERTED')
            else:
                stockDataRowAsList.append(stock.financialCurrency + ', now converted to USD')
            stockDataRowAsList.append(list(numpy.around(numpy.array(stock.conversionFactorToConvertToUSD), 4)))#round conversion factor to 4 decimal places

        if (math.isnan(stock.marketCap)):
            stockDataRowAsList.append('No data')
        else:
            stockDataRowAsList.append(stock.marketCap)

        if (math.isnan(stock.volume)):
            stockDataRowAsList.append('No data')
        else:
            stockDataRowAsList.append(stock.volume)

        if (math.isnan(stock.dividendRate)):
            stockDataRowAsList.append('N/A')
        else:
            stockDataRowAsList.append(stock.dividendRate)

        if (math.isnan(stock.dividendYield)):
            stockDataRowAsList.append('N/A')
        else:
            stockDataRowAsList.append(stock.dividendYield)

        noLastDateMessage = 'None within past year'
        if (len(stock.intersectionData['50dayMAOver200dayMA']['upCrossDates']) > 0):
            stockDataRowAsList.append(stock.intersectionData['50dayMAOver200dayMA']['upCrossDates'][-1])
        else:
            stockDataRowAsList.append(noLastDateMessage)

        if (len(stock.intersectionData['50dayMAOver200dayMA']['downCrossDates']) > 0):
            stockDataRowAsList.append(stock.intersectionData['50dayMAOver200dayMA']['downCrossDates'][-1])
        else:
            stockDataRowAsList.append(noLastDateMessage)

        if (len(stock.intersectionData['14dayMAOver200dayMA']['upCrossDates']) > 0):
            stockDataRowAsList.append(stock.intersectionData['14dayMAOver200dayMA']['upCrossDates'][-1])
        else:
            stockDataRowAsList.append(noLastDateMessage)

        if (len(stock.intersectionData['14dayMAOver200dayMA']['downCrossDates']) > 0):
            stockDataRowAsList.append(stock.intersectionData['14dayMAOver200dayMA']['downCrossDates'][-1])
        else:
            stockDataRowAsList.append(noLastDateMessage)

        if (len(stock.intersectionData['priceOver200dayMA']['upCrossDates']) > 0):
            stockDataRowAsList.append(stock.intersectionData['priceOver200dayMA']['upCrossDates'][-1])
        else:
            stockDataRowAsList.append(noLastDateMessage)

        if (len(stock.intersectionData['priceOver200dayMA']['downCrossDates']) > 0):
            stockDataRowAsList.append(stock.intersectionData['priceOver200dayMA']['downCrossDates'][-1])
        else:
            stockDataRowAsList.append(noLastDateMessage)






        ##
        ##ADD DATA FOR CURRENT ROW
        stockData2DListForDataFrame.append(stockDataRowAsList)




    ##
    ##SAVE EXCEL FILE

    #create data frame representing all stock data gathered
    stockData_DataFrame = pd.DataFrame(stockData2DListForDataFrame, columns=stockDataColumnTitlesForDataFrame)



    # if I want revenues, profits and margins all next to eachother, change the column order
    cols_stockData_DataFrameRef = list(stockData_DataFrame.columns.values)

    newQuarterlyColOrder = []
    newQuarterlyColOrder.extend(quarterlyDataColumnTitles[0::3])#revenues
    newQuarterlyColOrder.extend(quarterlyDataColumnTitles[1::3])#margins
    newQuarterlyColOrder.extend(quarterlyDataColumnTitles[2::3])#profits

    #numberOfColumnsBeforeQuarterlyCols = len(initialColumnTitles)
    newColsOrder = cols_stockData_DataFrameRef[:len(initialColumnTitles)]
    newColsOrder.extend(newQuarterlyColOrder)
    newColsOrderLength = len(newColsOrder)
    newColsOrder.extend(cols_stockData_DataFrameRef[newColsOrderLength:len(additionalColumnTitles)+newColsOrderLength])


    stockData_DataFrame = stockData_DataFrame[newColsOrder]




    #save excel file
    currentTimeAndDate = datetime.now(timezone('US/Eastern'))
    stockDataExcelFileName = 'stockData_'+dateOfInterest.strftime('%m-%d-%Y___') + currentTimeAndDate.strftime('generated_%m-%d-%Y_%I-%M%p_EST') + '.xlsx'
    stockData_DataFrame.to_excel(stockDataExcelFileName, index=False)





    ##
    ##FORMAT EXCEL FILE

    #apply formatting to the excel file
    #possibly a better way to open the file than this. If I set visible to true, a seperate workbook and new sheet opens up each time
    excel_app = xlwings.App(visible=False)
    excel_book = excel_app.books.open(stockDataExcelFileName)
    sht1 = excel_book.sheets['Sheet1']

    #should be a way to do this without iterating through. Understand range data type better
    for i in range(len(stockSymbolList)):
        sht1.cells(i + 2, 'A').number_format = '[Color10]+0.00\%;[Red]-0.0#\%'#sorting value, change % from march lows in this case
        sht1.cells(i + 2, 'C').number_format = '$0.00;-$0.00'#stock price
        sht1.cells(i + 2, 'D').number_format = '[Color10]+$0.00;[Red]-$0.00'#change in dollars
        sht1.cells(i + 2, 'E').number_format = '[Color10]+0.00\%;[Red]-0.00\%'#change in %
        sht1.cells(i + 2, 'F').number_format = '0.00;-0.00'#pe ratio

        # this formats data if it's in the all 4 Q rev, all 4 Q profit, all 4 Q margin order
        numberOfColumnsBeforeQuarterlyData = len(initialColumnTitles)
        for j in range(len(quarterlyDataColumnTitles)):
            if (j < len(sortedListOfAllQuarterNames)*2):
                sht1.cells(i + 2, j + 1 + numberOfColumnsBeforeQuarterlyData).number_format = '[Color10]0,00;[Red]-0,00'

            else:
                sht1.cells(i + 2, j + 1 + numberOfColumnsBeforeQuarterlyData).number_format = '[Color10]0.00\%;[Red]-0.00\%'

        '''
        #this formats data if it's in the rev profit margin order triplets for each quarter
        numberOfColumnsBeforeQuarterlyData = len(stockDataColumnTitlesForDataFrame) - len(quarterlyDataColumnTitles)
        for j in range(len(quarterlyDataColumnTitles)):
            if(((j+1)%3) == 0):#every 3rd quarterly data column is margins data, which needs different formatting
                sht1.cells(i + 2, j + 1 + numberOfColumnsBeforeQuarterlyData).number_format = '[Color10]0.00\%;[Red]-0.00\%'
            else:
                sht1.cells(i + 2, j+1+numberOfColumnsBeforeQuarterlyData).number_format = '[Color10]0,00;[Red]-0,00'
        '''

    #formating additional data, using range method
    numberOfcolumnsBeforeAdditionalData = len(initialColumnTitles)+len(quarterlyDataColumnTitles)

    numberColBeforeMarketCapInAddtlData = additionalColumnTitles.index(currentTimeAndDate.strftime('Latest Market cap (as of %m-%d-%Y)'))
    marketCapAndVolRange = sht1[ 1:(len(stockSymbolList)+1),(numberOfcolumnsBeforeAdditionalData+numberColBeforeMarketCapInAddtlData):(numberOfcolumnsBeforeAdditionalData+numberColBeforeMarketCapInAddtlData+2)]
    marketCapAndVolRange.number_format = '0,00;-0,00'

    numberColBeforeDividendInAddtlData = additionalColumnTitles.index(currentTimeAndDate.strftime('Latest Dividend (as of %m-%d-%Y)'))
    dividendRange = sht1[1:(len(stockSymbolList)+1),numberOfcolumnsBeforeAdditionalData+numberColBeforeDividendInAddtlData]
    dividendRange.number_format = '$0.00;-$0.00'
    dividendYieldRange = sht1[1:(len(stockSymbolList)+1),numberOfcolumnsBeforeAdditionalData+numberColBeforeDividendInAddtlData+1]
    dividendYieldRange.number_format = '0.00%;-0.00%'#since no \% used, % this converts data to percent out of 100 by multiplying by 100

    #alignment
    allDataRange = sht1[1:(len(stockSymbolList)+1), 0:len(stockDataColumnTitlesForDataFrame)]
    allDataRange.api.HorizontalAlignment = -4152  # left aligned: -4131 right aligned: -4152 center aligned: -4108
    sht1.range('B2:B' + str(len(stockSymbolList)+1)).api.HorizontalAlignment = -4131

    numberColBeforeFinCurConvInAddtlData = additionalColumnTitles.index('Conversion factors that were used to convert financial data to USD (left to right, latest to oldest quarter)')
    financialCurrencyConversionRange = sht1[1:(len(stockSymbolList)+1),numberOfcolumnsBeforeAdditionalData+numberColBeforeFinCurConvInAddtlData]
    financialCurrencyConversionRange.api.HorizontalAlignment = -4131

    #wrapping and column width/height
    wrapTextColumnHeaderRange = sht1[0, numberOfcolumnsBeforeAdditionalData+1:(numberOfcolumnsBeforeAdditionalData+6)]
    wrapTextColumnHeaderRange.api.WrapText = True
    sht1.autofit(axis="columns")
    sht1.autofit(axis="rows")

    #sht1.autofit()

    excel_book.save()
    excel_book.close()
    excel_app.quit()





    print('\n\n-------\n' + dateOfInterest.strftime('%m-%d-%Y') + ' stock report finished and saved' + '\n-------\n\n\n\n')

    return stockDataExcelFileName, graphPdfFileName

