from typing import List, Any

import pandas as pd
import numpy as np
import matplotlib.pyplot as pyplot
import matplotlib.dates as mdates
#%matplotlib inline
import seaborn
import datetime
from datetime import timedelta
from datetime import datetime
import numpy
from yahooquery import Ticker
import yahooquery
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.patheffects as pe


def getIntersectionDates(dataFrameToIntersectWith, intersectingDataFrame ,dates):
    df = pd.DataFrame(index=dates,columns=['Diff','Cross'])
    df['Diff'] = dataFrameToIntersectWith.close - intersectingDataFrame.close
    df['Cross'] = numpy.select([((df.Diff < 0) & (df.Diff.shift() > 0)), ((df.Diff > 0) & (df.Diff.shift() < 0))], ['Up', 'Down'], 'None')

    upCrossDataFrame = df.loc[df['Cross'] == 'Up']
    downCrossDataFrame = df.loc[df['Cross'] == 'Down']
    upCrossDates = list(upCrossDataFrame.index.values)
    downCrossDates: list[Any] = list(downCrossDataFrame.index.values)

    return upCrossDates, downCrossDates



def saveFiguresAsMultiPagePDF(figures, startDate, endDate, subPlotNrows, subPlotNcols):
    # Create the PdfPages object to which we will save the pages:
    # The with statement makes sure that the PdfPages object is closed properly at
    # the end of the block, even if an Exception occurs.
    pdfFileName = ('multipage_pdf_OfStockGraphs_' + startDate.strftime('%m-%d-%Y') + '_to_' + endDate.strftime('%m-%d-%Y') + '_' + str(subPlotNrows) + 'x' + str(subPlotNcols) +
                  datetime.today().strftime('__generated_%m-%d-%Y_%I-%M%p_EST') + '.pdf')
    with PdfPages(pdfFileName) as pdf:
        for fig in figures:
            #pdf.savefig(fig, orientation='landscape')
            pdf.savefig(fig)

        d = pdf.infodict()
        d['Title'] = 'Multipage PDF Of Stock Graphs'
        d['Author'] = u'Mike Ferguson \xe4nen'
        d['Subject'] = 'Graphing multiple matplotlib figures of Stock graphs showing golden crosses and other intersections'
        d['Keywords'] = 'stock graphs'
        d['CreationDate'] = datetime.today()
        d['ModDate'] = datetime.today()

        return pdfFileName



def graphStockDataAndReturnStockDataUpdatedWithIntersectionData(listofStockData, graphEndDate, graphStartDate = None):
    if graphStartDate == None:
        graphStartDate = graphEndDate - timedelta(weeks=52)#default graph is a 1 year graph

    seaborn.set(style='darkgrid', context='talk', palette='Dark2')
    my_year_month_fmt = mdates.DateFormatter('%m-%d-%y')

    SMALL_SIZE = 10
    MEDIUM_SIZE = 15
    BIGGER_SIZE = 20
    STOCK_SYMBOL_SIZE = 30
    #defaults only
    pyplot.rc('font', size=MEDIUM_SIZE)  # controls default text sizes, e.g. annotate text size
    pyplot.rc('axes', titlesize=STOCK_SYMBOL_SIZE)  # fontsize of the axes title, i.e. title of the individual subplots!
    pyplot.rc('axes', labelsize=SMALL_SIZE)  # fontsize of the x and y labels
    pyplot.rc('xtick', labelsize=5)  # fontsize of the x tick labels
    pyplot.rc('ytick', labelsize=SMALL_SIZE)  # fontsize of the y tick labels
    pyplot.rc('legend', fontsize=SMALL_SIZE)  # legend fontsize
    pyplot.rc('figure', titlesize=BIGGER_SIZE)  # fontsize of the figure title


    annotationTextColor = 'black'
    annotationTextOutlineColor = 'white'
    annotationTextOutlineLineWidth = 2
    defaultTextColor = 'black'
    pyplot.rc('text', color=defaultTextColor)

    #pyplot.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=None, hspace=None)
    pyplot.subplots_adjust(hspace=0.4)
    '''
    The parameter meanings( and suggested defaults) are:
    
    left = 0.125  # the left side of the subplots of the figure
    right = 0.9  # the right side of the subplots of the figure
    bottom = 0.1  # the bottom of the subplots of the figure
    top = 0.9  # the top of the subplots of the figure
    wspace = 0.2  # the amount of width reserved for blank space between subplots
    hspace = 0.2  # the amount of height reserved for white space between subplots
    '''
    subPlotNrows = 10
    subPlotNcols = 10
    annotationTextXShift = timedelta(days=5*subPlotNcols)#days=25 works well for 5x5 plot
    #pyplot.tight_layout()#sets tight layout for all figures or does it? doesn't work?
    figSize = ((11/(8.5/11))*2, 11*2)#width, height in inches
    #figSize = (8.5*2, (8.5/(11/8.5))*2)#width, height in inches
    #figSize = (8.5,11)
    figNum = 1
    #fig, axes = pyplot.subplots(nrows=subPlotNrows, ncols=subPlotNcols, figsize=figSize, subplot_kw=dict(box_aspect=1),sharex=False, sharey=False, constrained_layout=False)#, gridspec_kw={'width_ratios': [1]*subPlotNrows, 'height_ratios': [1]*subPlotNcols})#see gridspec documentation to understand what it does

    fig, axes = pyplot.subplots(nrows=subPlotNrows, ncols=subPlotNcols, figsize=figSize)
    fig.suptitle('Figure ' + str(figNum) + '- Date Range: ' + graphStartDate.strftime('%m-%d-%Y') + ' to ' + graphEndDate.strftime('%m-%d-%Y'), x=0, ha='left', y=0.9975)#y changes location of figure title in y direction, it's a percentage (0-1)
    fig.tight_layout()#sets tight layout for individual figure
    masterFigLegendLabels = []
    masterFigLegendHandles = []
    listOfFigures =[]
    listOfFigures.append(fig)

    numStocks = len(listofStockData)
    stockIterator = 0
    axesRow = 0
    axesCol = 0
    print(figNum)
    print(str(axesRow) + ',' + str(axesCol))
    for stock in listofStockData:

        yahooQueryStock = Ticker(stock.stockSymbol, status_forcelist=[404, 429, 500, 502, 503, 504], backoff_factor=10,timeout=None)
        data = yahooQueryStock.history(start=graphEndDate - timedelta(days=365 * 2), end=graphEndDate + timedelta(days=1))

        #this set of code just basically transforms the dataframe given by yahooquery's historical data into a more practical dataframe, gets rid of annoying tuple index
        data_closePrices = list(data.loc[:, 'close'].to_numpy())
        data_indexTuples = list(data.index.values)
        data_datesList = [i[1] for i in data_indexTuples]
        closePricesDataFrame = pd.DataFrame(data_closePrices, index=data_datesList, columns=['close'])


        movingAvg_14day = closePricesDataFrame.rolling(window=14).mean()
        movingAvg_50day = closePricesDataFrame.rolling(window=50).mean()
        movingAvg_200day = closePricesDataFrame.rolling(window=200).mean()

        listOfDataFramesToTestForIntersectionWith200DayMA = [closePricesDataFrame, movingAvg_14day, movingAvg_50day]
        numIntersectionsToTrack = len(listOfDataFramesToTestForIntersectionWith200DayMA)
        listOfUpCrossDatesForIntersectionsToTrack = []
        listOfDownCrossDatesForIntersectionsToTrack = []
        for i in range(numIntersectionsToTrack):
            upCrossDates, downCrossDates = getIntersectionDates(movingAvg_200day, listOfDataFramesToTestForIntersectionWith200DayMA[i], data_datesList)
            listOfUpCrossDatesForIntersectionsToTrack.append(upCrossDates)
            listOfDownCrossDatesForIntersectionsToTrack.append(downCrossDates)

        #create 2 level dictionary for storing intersection data. This will be returned and can be used to input data into excel sheet
        dictionaryKeysForIntersectionDataToTrack = ['priceOver200dayMA', '14dayMAOver200dayMA', '50dayMAOver200dayMA']
        intersectionDataToTrackDict = {}
        for i in range(numIntersectionsToTrack):
            intersectionDataToTrackDict.update( { dictionaryKeysForIntersectionDataToTrack[i] : {'upCrossDates':listOfUpCrossDatesForIntersectionsToTrack[i], 'downCrossDates':listOfDownCrossDatesForIntersectionsToTrack[i] } })
        stock.intersectionData = intersectionDataToTrackDict#this updates the stock objects sent into the function. I can return the modified objects at the end
        keysForIntersectionDataToGraph = ['priceOver200dayMA','14dayMAOver200dayMA', '50dayMAOver200dayMA']#list of dict keys corresponding to intersection data to graph


        ##
        ##PLOT BASE GRAPHS

        #get dates to plot based on desired graph start date.
        datesToPlot = []
        for d in data_datesList:
            if (d >= graphStartDate.date()):
                datesToPlot.append(d)
        start_date = datesToPlot[0]
        end_date = datesToPlot[-1]

        #get correct axes, way to access correct axes depends on number of rows and cols
        if(subPlotNrows > 1 and subPlotNcols > 1):
            ax = axes[axesRow, axesCol]#ax actually points to the axes object here. It's not a seperate variable
        elif(subPlotNrows > 1):
            ax = axes[axesRow]
        elif(subPlotNcols > 1):
            ax = axes[axesCol]
        else:
            ax = axes

        #plotting the initial graph without intersections, just price and moving averages
        ax.plot(datesToPlot, closePricesDataFrame.loc[start_date:end_date, 'close'], 'b', label='Price')
        ax.plot(datesToPlot, movingAvg_14day.loc[start_date:end_date, 'close'], 'g', label='14 day SMA')
        ax.plot(datesToPlot, movingAvg_50day.loc[start_date:end_date, 'close'], '--', color='tan', label='50 day SMA')
        ax.plot(datesToPlot, movingAvg_200day.loc[start_date:end_date, 'close'], color='pink', label='200 day SMA')



        ##
        ##PLOT INTERSECTIONS

        # get crossing dates within plotting range
        datesToPlot_as_set = set(datesToPlot)
        for key in keysForIntersectionDataToGraph:

            upCrossDatesWithinDesiredPlotDateRange = list(datesToPlot_as_set.intersection(intersectionDataToTrackDict[key]['upCrossDates']))
            downCrossDatesWithinDesiredPlotDateRange = list(datesToPlot_as_set.intersection(intersectionDataToTrackDict[key]['downCrossDates']))


            #plot intersections
            if(key == '50dayMAOver200dayMA'):
                ax.plot(upCrossDatesWithinDesiredPlotDateRange, movingAvg_200day.loc[upCrossDatesWithinDesiredPlotDateRange, 'close'], 'D', color='gold', label='Golden Cross')
                ax.plot(downCrossDatesWithinDesiredPlotDateRange, movingAvg_200day.loc[downCrossDatesWithinDesiredPlotDateRange, 'close'], 'D', color='red', label='Death Cross')
            elif (key == '14dayMAOver200dayMA'):
                ax.plot(upCrossDatesWithinDesiredPlotDateRange, movingAvg_200day.loc[upCrossDatesWithinDesiredPlotDateRange, 'close'], '*', color='yellow', label='14MA cross up over 200MA')
                ax.plot(downCrossDatesWithinDesiredPlotDateRange, movingAvg_200day.loc[downCrossDatesWithinDesiredPlotDateRange, 'close'], '*', color='red', label='14MA cross down under 200MA')
            elif (key == 'priceOver200dayMA'):
                ax.plot(upCrossDatesWithinDesiredPlotDateRange, movingAvg_200day.loc[upCrossDatesWithinDesiredPlotDateRange, 'close'], 'x', color='yellow', label='price cross up over 200MA')
                ax.plot(downCrossDatesWithinDesiredPlotDateRange, movingAvg_200day.loc[downCrossDatesWithinDesiredPlotDateRange, 'close'], 'x', color='red', label='price cross down under 200MA')

            #annotate select intersections
            if (key == '50dayMAOver200dayMA'):
                for upCrossDate in upCrossDatesWithinDesiredPlotDateRange:
                    upCrossDateY = movingAvg_200day.loc[upCrossDate, 'close']
                    ax.annotate(upCrossDate.strftime('%m-%d-%y'), xy=(upCrossDate,upCrossDateY),  xytext=(upCrossDate-annotationTextXShift, upCrossDateY), path_effects=[pe.withStroke(linewidth=annotationTextOutlineLineWidth, foreground=annotationTextOutlineColor)], color=annotationTextColor)
                for downCrossDate in downCrossDatesWithinDesiredPlotDateRange:
                    downCrossDateY = movingAvg_200day.loc[downCrossDate, 'close']
                    ax.annotate(downCrossDate.strftime('%m-%d-%y'), xy=(downCrossDate, downCrossDateY), xytext=(downCrossDate - annotationTextXShift, downCrossDateY), path_effects=[pe.withStroke(linewidth=annotationTextOutlineLineWidth, foreground=annotationTextOutlineColor)], color=annotationTextColor)

        #ax.legend(loc='best', fontsize='xx-small', frameon=True, framealpha=0.3, labelspacing=0.2)#fontsize can be set to very small number if not small enough
        #ax.legend(loc='best', fontsize=5, framealpha=0.3, markerscale=0.35)
        ax.set_ylabel('Closing price in USD ($)')
        #ax.set_xlabel(xlabel='Date (month-day-year)', fontsize=3)
        ax.set_title(label=stock.stockSymbol, pad=1)#xlim means graph x axis limited to those start and stop points. no empty space on sides
        ax.xaxis.set_major_formatter(my_year_month_fmt)
        # Ensure a major tick for desired interval
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=4))#mdates.WeekdayLocator(interval=2) means 2 weeks
        pyplot.setp(ax.get_xticklabels(), rotation='horizontal', position=(0,0.035))#position moves x tick labels in x and y dimension, rotation can be a number in degrees
        #update master handle and label lists, only add ones that don't exist
        handles, labels = ax.get_legend_handles_labels()
        for handle in handles:
            if handle not in masterFigLegendHandles:
                masterFigLegendHandles.append(handle)
        for label in labels:
            if label not in masterFigLegendLabels:
                masterFigLegendLabels.append(label)


        stockIterator += 1
        if(axesCol == (subPlotNcols-1)):
            if ((axesRow == (subPlotNrows-1)) and (stockIterator < numStocks)):
                #fig.suptitle('Stock: ' + stock.stockSymbol, fontsize=60)
                #fig.autofmt_xdate()  # makes dates slanted so not jumbled together


                figNum += 1
                axesRow = 0
                axesCol = 0
                fig, axes = pyplot.subplots(nrows=subPlotNrows, ncols=subPlotNcols, figsize=figSize)
                fig.suptitle('Figure ' + str(figNum) + '- Date Range: ' + graphStartDate.strftime('%m-%d-%Y') + ' to ' + graphEndDate.strftime('%m-%d-%Y'), x=0, ha='left', y=0.9975)  # y changes location of figure title in y direction, it's a percentage (0-1)
                fig.tight_layout()  # sets tight layout for individual figure
                listOfFigures.append(fig)
                print(figNum)
            else:
                axesRow += 1
                axesCol = 0
        else:
            axesCol += 1

        print(str(axesRow) + ',' + str(axesCol))




    for fig in listOfFigures:
        fig.legend(masterFigLegendHandles, masterFigLegendLabels, loc='upper right', ncol = len(masterFigLegendHandles), labelspacing=0)

    #pyplot.show()
    graphPdfFilename = saveFiguresAsMultiPagePDF(listOfFigures,graphStartDate,graphEndDate,subPlotNrows,subPlotNcols)
    for fig in listOfFigures:
        pyplot.close(fig)
    pyplot.close('all')  # sometimes a blank figure 1 window is left over, why? Google this. Maybe this will fix it





    return graphPdfFilename, listofStockData




