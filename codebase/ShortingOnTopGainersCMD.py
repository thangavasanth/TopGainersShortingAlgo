#Shorting on Top gainers Everday 
#**Program Starts here**

from ib_insync import *
import pandas as pd
import os
import xlwings as xw
from datetime import datetime, timedelta
from datetime import date
from decimal import *
import schedule
import time
import math
import csv  
import threading
import time
import numpy as np
from numpy import nan
import csv
import sys
import asyncio
import nest_asyncio
import random
nest_asyncio.apply()
util.startLoop()
util.patchAsyncio()
# script_path = os.path.realpath(sys.argv[0])
# print("Executing: ", script_path)


try:
# if 1==1:    
    maxTry = 1
    now = datetime.now()
        
    # excel_file = r'C:\Vasanth\xlWingsAlgo\EncryptedFiles\Yazhi_Algo_Trading.xlsm'
    # sheet_name = 'Top Gainers Shorting'
    excel_file = r'[Yaali_Algo_Trading_Full_FilePath]'
    sheet_name = '[Yaali_Algo_Sheet_Name]'
    # excel_file =  r'' + "[{ExcelFileLocation}]"
    # sheet_name = "[{ExcelSheetName}]"
    # print(excel_file)
    # print(sheet_name)
    # ib.sleep(20)
    start_row = 9
    
    wb = xw.Book(excel_file)
    sheet = wb.sheets[sheet_name]
    # accountIBKR= sheet.range(f'C2').value 
            
    rootPath = sheet.range('M2').value
    RunType = sheet.range('I7').value                
    MaxPositionsInPortfolio = sheet.range(f'K2').value 
    maxCashPerStock = sheet.range(f'K3').value 
    profitTakerPercent =  sheet.range(f'K4').value 
    stopLossPercent =  sheet.range(f'K5').value 
    HowManyTickersForEachClient =  sheet.range(f'K6').value 
    average_percent = sheet.range('I4').value
    entryPercentage = sheet.range(f'E3').value 
    avgVolumeAbove =  sheet.range(f'E4').value 
    priceBelow =  sheet.range(f'E6').value 
    priceAbove =  sheet.range(f'E7').value
    
    #print(priceAbove)
    RSIAbove =  sheet.range(f'G2').value 
    RSIAboveifLastPriceLessOneDollar =sheet.range(f'G3').value    
    RSIDifferenceAbove = sheet.range(f'G4').value        
    changePercentIn15Min = sheet.range(f'G6').value 
    HighlighExcelCells = sheet.range(f'G7').value 
    # print(rootPath)
    data_folder = f'{rootPath}/datafiles'
    codebase_folder =  f'{rootPath}/codebase'
    log_folder =  f'{rootPath}/logs'
    archive_folder = f'{data_folder}/archive'
    today = datetime.now().strftime('%Y_%m_%d')
    # File Names and Paths
   
    if not os.path.exists(log_folder):
        os.makedirs(log_folder)
        
    
    
    LogFileName = "TopGainersLog.csv"
    LogFileName = os.path.join(log_folder, LogFileName) 
    if os.path.exists(LogFileName):
        #print (LogFileName)
        ArchiveLogFileName = log_folder + "\\TopGainersLog_" + now.strftime("%Y_%m_%d_%H_%M_%S") + ".csv"
        #BackupFileName = datetime.now.strftime("%Y_%m_%d_, %H_%M_%S") + ".csv"
        os.rename(LogFileName, ArchiveLogFileName)
    ib = IB()
    with open(LogFileName, 'a', encoding='UTF8', newline='') as f2: # or 'wb' if on python2
        writer = csv.writer(f2)
        writer.writerow(['Time','Round', 'Rank', 'Stock','Prev Day Close', 'Last Price', 'Bid Price', 'Ask Price', 'High Price', 'Low Price' ,'Change Percentage', 'RSI Now', 'RSI 5 min Back','RSI Difference', 'Position', 'Condition Satisfied']) # replace with your custom column header
    # print("Log File Created:", LogFileName)
    def ShortingOnTopGainers(Nthtime, accountIBKR,  ClientID):
         # try:
         if 1==1:
            show_type = sheet.range('J7').value    
            if show_type == 'Exit':
                # print ("Exitting 1")
                sheet.range(f'B9:V24').color = None
                ib.disconnect()
                sys.exit()
                        
            RunType = sheet.range('I7').value                
            MaxPositionsInPortfolio = sheet.range(f'K2').value 
            maxCashPerStock = sheet.range(f'K3').value 
            profitTakerPercent =  sheet.range(f'K4').value 
            stopLossPercent =  sheet.range(f'K5').value 
            HowManyTickersForEachClient =  sheet.range(f'K6').value 
            average_percent = sheet.range('I4').value
            entryPercentage = sheet.range(f'E3').value 
            avgVolumeAbove =  sheet.range(f'E4').value 
            priceBelow =  sheet.range(f'E5').value 
            priceAbove =  sheet.range(f'E6').value 
            
            #print(priceAbove)
            RSIAbove =  sheet.range(f'G2').value 
            RSIAboveifLastPriceLessOneDollar =sheet.range(f'G3').value    
            RSIDifferenceAbove = sheet.range(f'G4').value  
            changePercentIn15Min = sheet.range(f'G6').value 
            HighlighExcelCells = sheet.range(f'G7').value 
            StockCategory = ""
            
            FromRow = int(start_row + ((ClientID * HowManyTickersForEachClient) - HowManyTickersForEachClient))
            ToRow = int(start_row + ClientID * HowManyTickersForEachClient-1)
            scanData = sheet.range( f'B{FromRow}:B{ToRow}').value
            if scanData:
                scanData = [scanData] if not isinstance(scanData, list) else scanData
                scanData = [cell for cell in scanData if cell is not None]

            # print (scanData) 
            print("Account Summary of IBKR Account: ", accountIBKR, "Cliend ID:", ClientID)
            print("Running in...", show_type)
            print("Monitoring For Tickers:", scanData)
            if show_type == 'Show in Console':
                # print("here 1")
                print("_______________Shorting On Top Gainers__________________________________________________________________________________Entry Criteria_______________")
                print("Entry Percentage Above: {:10s}                            Price Below: {:10s}".format(str(entryPercentage),str(priceBelow)))
                print("   RSI Indicator Above: {:10s}                   Average Volumn Above: {:10s}".format(str(RSIAbove),str(avgVolumeAbove)))
                print("   RSI Indicator Latest should be Above RSI Indicator 5 Minutes Back")
                print("_____________________________________________________________________________________________________________________________________________________")
            i = 0
            for Stocksymbol in scanData:
                # os.system('cls' if os.name == 'nt' else 'clear')   
                alreadyOpenOrdersExists = False
                # Stocksymbol = scan
                # print(Stocksymbol)
                row = FromRow + (i) 
                MonitoringStatus = sheet.range('C' + str(row)).value 
                if show_type == 'Show in Excel':
                    print("Monitoring for Tikcer:" , Stocksymbol, "              Excel  Row: ", row)
                    sheet.range('B' + str(row)).color = "#a9d08e"
                    # sheet.range('C' + str(row)).color = "#a9d08e"
                    ib.sleep(0.1)
                    sheet.range('B' + str(row)).color = None
                    # sheet.range('C' + str(row)).color = None
                    if not MonitoringStatus or MonitoringStatus=="Monitoring":
                        sheet.range('C' + str(row)).color = None
                        PrintMessage(show_type,row,'C', "Monitoring","#a9d08e") 
                        
                        
                
                # i = i+1
                # continue
                contract=Stock(Stocksymbol, 'SMART', 'USD')
                dfContractDetails = ib.reqContractDetails(contract)
                StockIndustry = dfContractDetails[0].industry
                StockCategory = dfContractDetails[0].category
                if show_type == 'Show in Console':
                    print("_______Round:{:3}______________________________________________________________________Rank :{:2}________Stock:{}    Time: {}".format(str(Nthtime),str(i+1),Stocksymbol,datetime.now()))
                    print("Industry: {}                                  Category:{}".format(StockIndustry,StockCategory))
                #     #if dfContractDetails[0].industry == 'Energy' or dfContractDetails[0].category == 'Energy' or Stocksymbol == "IMPP":
                contract=Stock(Stocksymbol, 'SMART', 'USD')
                mktData = ib.reqMktData(contract, "165,236,233,318", False, False, [])
                ib.sleep(3)
                
                
                halted = mktData.halted #Volatility halt = 2
                reTry = 0
                while (not (mktData.last >0  and mktData.close > 0 and mktData.bid > 0 and mktData.ask > 0) and not reTry> maxTry):
                # while (not (mktData.last >0) and not reTry> maxTry and not halted>0):
                    reTry = reTry+1 
                    sheet.range('C' + str(row)).options(index=False, header=False).value  =   "Re-Trying Mkt Data"
                    mktData = ib.reqMktData(contract, "165,236,233,318", False, False, [])
                    ib.sleep(3+reTry)
                    halted = mktData.halted
                     
                    # sheet.range('C' + str(row)).options(index=False, header=False).value  = str(reTry+1) +   " Re-try Market Data" 
                if reTry >= maxTry:
                    sheet.range('C' + str(row)).options(index=False, header=False).value  =   "Monitoring"
                    i = i+1
                    continue
                # PrintMessage(show_type,row,'C', halted, None)
                prvClosePrice = mktData.close
                lastPrice = mktData.last
                bidPrice = mktData.bid
                askPrice = mktData.ask
                highPrice = mktData.high
                lowPrice = mktData.low
                
                shortableShares = getattr(mktData, 'shortableShares', None)
                if shortableShares is  None or  np.isnan(shortableShares):
                   shortableShares = 0
                

                changePercent = int(round(((bidPrice - prvClosePrice)/prvClosePrice)*100.00,0))
                changePercentBasedBidPrice = int(round(((bidPrice - prvClosePrice)/prvClosePrice)*100.00,0))
                
                if (changePercent<entryPercentage-10) and not halted>0 and  not (MonitoringStatus in ("Stock Already in Portfolio", 'Evaluating To Average & Exit', 'Trying to Average & Exit')):
                    ib.reqAllOpenOrders()
                    for trade in ib.openTrades():
                        if (trade.contract.localSymbol == Stocksymbol or trade.contract.symbol == Stocksymbol) and trade.order.action == "SELL":
                            ib.cancelOrder(trade.order)
                            
                    sheet.range(f'B{row}:Z{row}').value = None
                    sheet.range(f'B{row}:Z{row}').color = None
                    ib.disconnect()
                    sys.exit()
                
                histData = ib.reqHistoricalData(contract,'','900 S','30 secs','ADJUSTED_LAST',1,1,0,[])
                ib.sleep(1)
                dfUtilhisData = util.df(histData)
                if(type(dfUtilhisData) == type(None)):
                    i=i+1
                    continue
                if not dfUtilhisData.empty:
                    dfRSI = RSI(dfUtilhisData,14)
                
                
                RSIndiacator5MinBack = round(dfRSI.values[len(dfRSI)-7],0)
                for x in range(22, len(dfRSI)-1 ):
                    RSIndiacatorinRange = round(dfRSI.values[x],0)
                    if RSIndiacator5MinBack > RSIndiacatorinRange:
                        RSIndiacator5MinBack = RSIndiacatorinRange
    
                #RSIndiacator5MinBack = round(dfRSI.values[len(dfRSI)-7],0)
                RSIIndicatorLatest =  round(dfRSI.values[len(dfRSI)-1],0)
                RSIDifference = RSIIndicatorLatest - RSIndiacator5MinBack
                
                lowest_value_15min = dfUtilhisData['low'].min()
                percentage_change_in_15_min = ((lastPrice - lowest_value_15min) / lowest_value_15min) * 100
                
       
                avgVolumne = mktData.avVolume*100
                todayVolumne =  mktData.volume*100
                avgVolumneInMillions = ""
                todayVolumneInMillions = ""
                if avgVolumne >= 1000000:
                    avgVolumneInMillions = str(round(avgVolumne/1000000,2)) + "M"
                else:
                    avgVolumneInMillions = avgVolumne
                if todayVolumne >= 1000000:
                    todayVolumneInMillions = str(round(todayVolumne/1000000,2)) + "M"
                else:
                    todayVolumneInMillions = todayVolumne
                    
                
                
                if show_type in ['Show in Excel', 'Show in Console']:
                    # PrintMessage(show_type,row,'D', prvClosePrice, None,HighlighExcelCells)
                    # PrintMessage(show_type,row,'D', "New Program", None)
                    PrintMessage(show_type,row,'D', lastPrice, None,HighlighExcelCells)
                    PrintMessage(show_type,row,'E', changePercent  , None, HighlighExcelCells)
                 
                    PrintMessage(show_type,row,'F', avgVolumneInMillions, None,HighlighExcelCells)
                    PrintMessage(show_type,row,'G', todayVolumneInMillions, None,HighlighExcelCells)
                    # PrintMessage(show_type,row,'I', shortableShares, None,HighlighExcelCells)
                    PrintMessage(show_type,row,'H', shortableShares, None,HighlighExcelCells)
                    PrintMessage(show_type,row,'I', RSIIndicatorLatest, None, HighlighExcelCells)
                    PrintMessage(show_type,row,'J', RSIndiacator5MinBack, None,HighlighExcelCells)
                    PrintMessage(show_type,row,'K', RSIDifference, None, HighlighExcelCells)
                    
                    PrintMessage(show_type,row,'L', percentage_change_in_15_min, None, HighlighExcelCells)
                    PrintMessage(show_type,row,'M', lowest_value_15min, None, HighlighExcelCells)
                    
                dfPortfolioPositionsUtil = util.df(ib.reqPositions())
                if str(dfPortfolioPositionsUtil) != 'None':
                    dfPortfolioPositionsUtil = dfPortfolioPositionsUtil.drop(dfPortfolioPositionsUtil.loc[dfPortfolioPositionsUtil['position'] == 0].index)
                    if len(dfPortfolioPositionsUtil) >= MaxPositionsInPortfolio:
                        PrintMessage(show_type,row,'C',"Portfolio Reached Max Stocks","#a9d08e")
                        return
                    alreadyExistsinPortfolio = False
                    for Portstock in dfPortfolioPositionsUtil.values:
                        if Portstock[1].symbol == Stocksymbol: # or Portstock[0] != accountIBKR:
                            alreadyExistsinPortfolio = True
                            break
                   
                    if alreadyExistsinPortfolio == True:
                        if not (MonitoringStatus in ("Stock Already in Portfolio", 'Evaluating To Average & Exit', 'Trying to Average & Exit')):
                            PrintMessage(show_type,row,'C',"Stock Already in Portfolio",None)
                        position_details = get_current_position(Stocksymbol)
                        if position_details:
                            position_size = position_details['Position Size']
                            avg_price = position_details['Average Cost']
                            base_price = avg_price * position_size
                            market_price = lastPrice * position_size
                            unrealized_pnl = (base_price - market_price)*-1
                            pnl_percentage = ((unrealized_pnl / base_price) * 100)*-1
                            
                            PrintMessage(show_type,row,'N', position_size, None,HighlighExcelCells)
                            PrintMessage(show_type,row,'O', unrealized_pnl, None, HighlighExcelCells)
                            PrintMessage(show_type,row,'P', pnl_percentage, None, HighlighExcelCells)
                            
                            # if pnl_percentage >0:
                            #      sheet.range('C' + str(row)).color = "#8fce00"
                            #      sheet.range('S' + str(row)).color = "#8fce00"
                            #      sheet.range('T' + str(row)).color = "#8fce00"
                            # elif pnl_percentage<=0:
                            #       sheet.range('C' + str(row)).color = "#FF97BA" 
                            #       sheet.range('S' + str(row)).color = "#FF97BA" 
                            #       sheet.range('T' + str(row)).color = "#FF97BA" 
                            
                            
                                 
                            if show_type in ['Show in Excel'] and HighlighExcelCells == "Yes":
                                newAverage_Price = lastPrice * average_percent
                                if newAverage_Price - lastPrice != 0:  # Check to prevent division by zero
                                    HowManyMorePostionsToShortAgain = (base_price - (newAverage_Price * position_size)) / (newAverage_Price - lastPrice)
                                else:
                                    HowManyMorePostionsToShortAgain = 0
                                    
                                HowMuchMoreAmountToShortAgain = newAverage_Price * HowManyMorePostionsToShortAgain
                                new_positionsTotal = position_size + HowManyMorePostionsToShortAgain
                                newbasePrice = new_positionsTotal * newAverage_Price
                                newmarketPrice = new_positionsTotal * lastPrice
                                
                                # PrintMessage(show_type,row,'O', position_size, None, HighlighExcelCells)
                                PrintMessage(show_type,row,'Q', avg_price, None, HighlighExcelCells)
                                PrintMessage(show_type,row,'R', base_price, None, HighlighExcelCells)
                                PrintMessage(show_type,row,'S', market_price, None, HighlighExcelCells)
                                # PrintMessage(show_type,row,'S', unrealized_pnl, None, HighlighExcelCells)
                                # PrintMessage(show_type,row,'T', pnl_percentage, None, HighlighExcelCells)
                                PrintMessage(show_type,row,'T', newAverage_Price, None, HighlighExcelCells)
                                PrintMessage(show_type,row,'U', HowManyMorePostionsToShortAgain, None, HighlighExcelCells)
                                PrintMessage(show_type,row,'V', HowMuchMoreAmountToShortAgain, None, HighlighExcelCells)
                                PrintMessage(show_type,row,'X', newbasePrice, None, HighlighExcelCells)
                                PrintMessage(show_type,row,'Y', newmarketPrice, None, HighlighExcelCells)
                            
                            
                            #     if pnl_percentage >0:
                            #         sheet.range('C' + str(row) + ":Z" + str(row) ).color = "#8fce00"
                            #     elif pnl_percentage<0:
                            #          sheet.range('C' + str(row) + ":Z" + str(row) ).color = "#FF97BA" 
                            #     else:
                            #         sheet.range('C' + str(row) + ":Z" + str(row) ).color = None
                            # else:
                            #     if pnl_percentage >0:
                            #          sheet.range('C' + str(row)).color = "#8fce00"
                            #     elif pnl_percentage<0:
                            #           sheet.range('C' + str(row)).color = "#FF97BA" 
                            #     else:
                            #          sheet.range('C' + str(row)).color = None
                        i = i+1
                        continue
                ib.reqAllOpenOrders()
                for trade in ib.openTrades():
                    if (trade.contract.localSymbol == Stocksymbol or trade.contract.symbol == Stocksymbol) and trade.order.action == "SELL":
                        # if trade.order.orderType == "LMT" and askPrice < trade.order.lmtPrice:
                        if trade.order.orderType == "LMT" and (askPrice < trade.order.lmtPrice or trade.orderStatus == "PreSubmitted"):
                            ib.cancelOrder(trade.order)
                        else:
                            alreadyOpenOrdersExists = True
                            # PrintMessage(show_type,row,'C', "Open Order Exists","#a9d08e")
                            break
                if alreadyOpenOrdersExists == True:
                    print("alreadyOpenOrdersExists")
                    PrintMessage(show_type,row,'C', "Open Order Exists","#a9d08e")
                    i = i+1
                    continue
                if halted == 2 or halted == 1:
                        shortOrder = Order()
                        shortOrder.account = accountIBKR
                        shortOrder.action = "SELL"
                        shortOrder.orderType = "MKT"
                        shortOrder.totalQuantity = int(maxCashPerStock/mktData.last)
                        #shortOrder.lmtPrice = limitPrice
                        #shortOrder.outsideRth = True
                        shortOrder.usePriceMgmtAlgo = True
                        shortOrder.transmit = True
                        if RunType == "Live":
                            dfshortOrder = ib.placeOrder(contract,shortOrder)
                        PrintMessage(show_type,row,'C', "Halted. Placed Short Order","#a9d08e")
                        i = i+1
                        continue
                    
                
                
                

                # if show_type in ['Show in Excel', 'Show in Console']:
                PrintMessage(show_type,row,'C', "Monitoring","#a9d08e") 
                sheet.range('N' + str(row) + ":P" + str(row)).clear_contents()
                sheet.range('N' + str(row) + ":P" + str(row)).color = None
                if show_type in ['Show in Console']:
                    ib.sleep(10)
                
                if lastPrice < 0.70:
                    RSIIndicatorAbove = RSIAboveifLastPriceLessOneDollar
                else:
                    RSIIndicatorAbove = RSIAbove

                quantity = int(maxCashPerStock/bidPrice)
                limitPrice = bidPrice
                decimalPlaces = len(str(limitPrice).split(".")[1])    
                ProfitlimitPrice = round(limitPrice - (profitTakerPercent/100)*limitPrice, decimalPlaces)
                stopLossPrice = round(limitPrice +  (stopLossPercent/100*limitPrice), decimalPlaces)

                ConditionSatisfied = 'No'
                if ((changePercent >= entryPercentage    #18
                     and ((RSIIndicatorLatest >  RSIndiacator5MinBack and RSIIndicatorLatest >=  RSIIndicatorAbove ) or (RSIDifference > RSIDifferenceAbove) )  
                      )):
                        if show_type == 'Show in Excel':
                            sheet.range('C' + str(row)).options(index=False, header=False).value  =   "Placed Short Order"
                            # sheet.range('D' + str(row) + ":T" + str(row) ).color = "#8fce00"
                            ib.sleep(0.5)
                        if show_type == 'Show in Console':
                            print("     ******Condition Satisfied******")
                        
                        if shortableShares <= 0:
                            maxCashPerStock = maxCashPerStock/2
                            
                        if (((RSIIndicatorLatest >=  95 ) or (RSIDifference > 75)) and (changePercent >= 60 )):
                            quantity = int((maxCashPerStock + maxCashPerStock*0.7)/bidPrice)
                        elif (((RSIIndicatorLatest >=  90 ) or (RSIDifference > 70)) and (changePercent >= 45)):
                            quantity = int((maxCashPerStock + maxCashPerStock*0.5)/bidPrice)
                        elif (((RSIIndicatorLatest >=  85 ) or (RSIDifference > 65)) and (changePercent >= 45)):
                            quantity = int((maxCashPerStock + maxCashPerStock*0.3)/bidPrice)
                        elif (((RSIIndicatorLatest >=  80 ) or (RSIDifference > 60)) and (changePercent >= 45 )):
                            quantity = int((maxCashPerStock + maxCashPerStock*0.2)/bidPrice)
                        elif (((RSIIndicatorLatest >=  75 ) or (RSIDifference > 55)) and (changePercent >= 45)):
                            quantity = int((maxCashPerStock + maxCashPerStock*0.1)/bidPrice)
                        else:
                            quantity = int(maxCashPerStock/bidPrice)
                        
                        # quantity = int(maxCashPerStock/bidPrice)
                        ConditionSatisfied = 'Yes'
                        parentOrder = ib.bracketOrder("SELL",quantity,limitPrice,ProfitlimitPrice,stopLossPrice)
                        parentOrder.parent.account = accountIBKR
                        parentOrder.parent.action = "SELL"
                        parentOrder.parent.orderType = "LMT"
                        parentOrder.parent.outsideRth = True
                        parentOrder.parent.usePriceMgmtAlgo = True
                        # parentOrder.parent.transmit = True
                        if RunType == "Live":
                            dfparentOrder = ib.placeOrder(contract,parentOrder.parent)
                        time.sleep(1)
                       
                        parentOrder.takeProfit.account= accountIBKR
                        parentOrder.takeProfit.orderType = "LMT"
                        #parentOrder.takeProfitr.orderId = parentOrder.parent.orderId +1
                        #parentOrder.takeProfit.parentId = parentOrder.parent.orderId
                        parentOrder.takeProfit.outsideRth = True
                        parentOrder.takeProfit.usePriceMgmtAlgo = True
                        parentOrder.takeProfit.tif = "GTC"
                        parentOrder.takeProfit.transmit = True
                        if RunType == "Live":           
                            dfprofitTakerOrder = ib.placeOrder(contract,parentOrder.takeProfit)
                        time.sleep(1)
                       
                        # #parentOrder.stopLoss.orderId = parentOrder.parent.orderId +2
                        # #parentOrder.stopLoss.parentId = parentOrder.parent.orderId
                        # parentOrder.stopLoss.account = accountIBKR
                        # parentOrder.stopLoss.orderType = "STP"
                        # #parentOrder.stopLoss.totalQuantity = quantity
                        # #parentOrder.stopLoss.lmtPrice = str(round(stopLossPrice - (10/100)*stopLossPrice, decimalPlaces))
                        # #parentOrder.stopLoss.auxPrice = str(round(stopLossPrice, decimalPlaces))
                        # parentOrder.stopLoss.lmtPrice = str(round(stopLossPrice - (10/100)*stopLossPrice, 2))
                        # parentOrder.stopLoss.auxPrice = str(round(stopLossPrice, 2))
                        # parentOrder.stopLoss.tif = "GTC"
                        # parentOrder.stopLoss.outsideRth = True
                        # parentOrder.stopLoss.usePriceMgmtAlgo = True
                        # parentOrder.stopLoss.transmit = True
                        # #if RunType == "Live":
                        #     # dfprofitTakerOrder = ib.placeOrder(contract,parentOrder.stopLoss)
                        time.sleep(1)
                        if show_type == 'Show in Console':
                            print("          Placed Short Order for Stock:{:6s} Qty:{:6s} LMT Price:{:8s} Amount:{:8s} Bid Price: {}"
                                    .format(Stocksymbol, str(int(quantity)),str(limitPrice),str(round(limitPrice*quantity,2)),bidPrice))
                else:
                    if show_type == 'Show in Console':
                        print("     XXXXX Condition NOT Satisfied XXXXX")
                ##print("_____Stock No: {}_______________________________________________________________________________________________Round:{}__________".format(i+1,Nthtime))
                with open('TopGainersLog.csv', 'a', encoding='UTF8', newline='') as f2: # or 'wb' if on python2
                    writer = csv.writer(f2)
                    #writer.writerow(['Time','Round', 'Stock', 'Change Percentage', 'RSI Now', 'RSI 5 min Back', 'Position', 'Limit Price', 'Condition Satisfied']) # replace with your custom column header
       
                    logdata = [datetime.now(),Nthtime,i, Stocksymbol, prvClosePrice,lastPrice, bidPrice, askPrice,highPrice,lowPrice, changePercent,RSIIndicatorLatest, RSIndiacator5MinBack,RSIDifference, quantity, ConditionSatisfied]
                    writer.writerow(logdata)
                i=i+1
         # except Exception as ex:
         #     print("Error 2:{}".format(str(ex)))
         #     # ib.disconnect()
         #     # time.sleep(5) 
    
    def PrintMessage(show_type, XLRowNumber, XLColumnCharacter, Value, Color=None,HighlighExcelCells="Blink" ):
        if (show_type == 'Show in Excel' and XLRowNumber > 0 and XLColumnCharacter != "") or (XLColumnCharacter == 'C'):
            cell = sheet.range(f'{XLColumnCharacter}{XLRowNumber}')
            old_value = cell.value
            old_color = cell.color
            
            if not Value:
                Value = 0
                
            if XLColumnCharacter in('E', 'P'):
                cell.value = str(Value) + "%"
            else:
                cell.value = Value 
            
            if XLColumnCharacter in ['B', 'C']:
                sheet.range(f'{XLColumnCharacter}{XLRowNumber}').color = Color
                ib.sleep(0.5)
                sheet.range(f'{XLColumnCharacter}{XLRowNumber}').color = old_color
            
            elif HighlighExcelCells.startswith("Yes"):
                # if XLColumnCharacter in ('E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'N', 'M') and sheet.range(f'C{XLRowNumber}').value != "Stock Already in Portfolio":
                if XLColumnCharacter in ('D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'N', 'M', 'O', 'P'):
                    RSINow = sheet.range(f'J{XLRowNumber}').value
                    if not RSINow:
                        RSINow = 0
                    if ((XLColumnCharacter == 'D' and Value <= priceBelow and Value >= priceAbove) 
                        or (XLColumnCharacter == 'E' and Value >= entryPercentage)
                        or (XLColumnCharacter in ('F', 'G') )
                        or (XLColumnCharacter == 'H' and Value >0)
                        or (XLColumnCharacter == 'I' and (Value >= RSIAbove or Value == 100))
                        or (XLColumnCharacter == 'J' and ((RSINow >= Value) or RSINow==100))
                        or (XLColumnCharacter == 'K' and Value >= RSIDifferenceAbove) 
                        or  (XLColumnCharacter == 'L' and Value >changePercentIn15Min)
                        
                        ):
                        cell.color = None
                        ib.sleep(0.2)
                        cell.color = "#A9D08E"
                        
                    elif (XLColumnCharacter in( 'N','O', 'P') and Value >0):
                        sheet.range('C' + str(XLRowNumber)).color = None
                        cell.color = None
                        ib.sleep(0.2)
                        cell.color = "#A9D08E"
                        sheet.range('C' + str(XLRowNumber)).color = "#A9D08E"
                        
                    elif (XLColumnCharacter in( 'N','O', 'P') and Value <=0):
                        sheet.range('C' + str(XLRowNumber)).color = None
                        cell.color = None
                        ib.sleep(0.2)
                        cell.color = "#FF97BA"
                        sheet.range('C' + str(XLRowNumber)).color = "#FF97BA"
                    else:
                        cell.color = None
                        ib.sleep(0.2)
                        cell.color = "#f4b084"
    
                else:
                    if Color:
                        cell.color = None
                        ib.sleep(0.2)
                        cell.color = Color
                          # Reset the color if a specific color was set and blink is not active
                    # else:
            elif HighlighExcelCells == "Blink":
                if old_value is not None and isinstance(old_value, (int, float)) and isinstance(Value, (int, float)):
                    if Value >= old_value:
                        cell.color = "#8fce00"
                    elif Value < old_value:
                        cell.color = "#f4b084"  # Red
                    ib.sleep(0.2)
                    cell.color = old_color  # Reset the color after the blink effect
            else: 
                cell.color = None 
                ib.sleep(0.2)
                cell.color = Color
            
               
                # #     cell.color = None
        elif show_type == 'Show in Console':
            fieldName = sheet.range(f'{XLColumnCharacter}{8}').value
            print(f"{fieldName:<30}: {Value}")
            
            
            
    def get_current_position(symbol):
        positions = ib.positions()
        for position in positions:
            if position.contract.symbol == symbol:
                position_details = {
                    'Symbol': position.contract.symbol,
                    'Position Size': position.position,
                    'Average Cost': position.avgCost,
                    'Account': position.account
                }
                return position_details
        
        # print(f"No position found for {symbol}.")
        return  None
    
    def RSI(DF,n=14):
         #     # sys.exit()
    #@ray.remote
        "function to calculate RSI"
        df = DF.copy()
        df['delta']=df['close'] - df['close'].shift(1)
        df['gain']=np.where(df['delta']>=0,df['delta'],0)
        df['loss']=np.where(df['delta']<0,abs(df['delta']),0)
        avg_gain = []
        avg_loss = []
        gain = df['gain'].tolist()
        loss = df['loss'].tolist()
        for i in range(len(df)):
            if i < n:
                avg_gain.append(np.NaN)
                avg_loss.append(np.NaN)
            elif i == n:
                avg_gain.append(df['gain'].rolling(n).mean()[n])
                avg_loss.append(df['loss'].rolling(n).mean()[n])
            elif i > n:
                avg_gain.append(((n-1)*avg_gain[i-1] + gain[i])/n)
                avg_loss.append(((n-1)*avg_loss[i-1] + loss[i])/n)
        df['avg_gain']=np.array(avg_gain)
        df['avg_loss']=np.array(avg_loss)
        df['RS'] = df['avg_gain']/df['avg_loss']
        df['RSI'] = 100 - (100/(1+df['RS']))
        return df['RSI']
    
    def MACD(DF,a=12,b=26,c=9):
        """function to calculate MACD
           typical values a(fast moving average) = 12;
                          b(slow moving average) =26;
                          c(signal line ma window) =9"""
        df = DF.copy()
        df["MA_Fast"]=df["close"].ewm(span=a,min_periods=a).mean()
        df["MA_Slow"]=df["close"].ewm(span=b,min_periods=b).mean()
        df["MACD"]=df["MA_Fast"]-df["MA_Slow"]
        df["Signal"]=df["MACD"].ewm(span=c,min_periods=c).mean()
        return df
    
    def STOCHOSCLTR(DF,a=20,b=3):
        """function to calculate Stochastics
           a = lookback period
           b = moving average window for %D"""
        df = DF.copy()
        df['C-L'] = df['close'] - df['low'].rolling(a).min()
        df['H-L'] = df['high'].rolling(a).max() - df['low'].rolling(a).min()
        df['%K'] = df['C-L']/df['H-L']*100
        #df['%D'] = df['%K'].ewm(span=b,min_periods=b).mean()
        return df['%K'].rolling(b).mean()
    
    def ATR(DF,n=20):
        "function to calculate True Range and Average True Range"
        df = DF.copy()
        df['H-L']=abs(df['high']-df['low'])
        df['H-PC']=abs(df['high']-df['close'].shift(1))
        df['L-PC']=abs(df['low']-df['close'].shift(1))
        df['TR']=df[['H-L','H-PC','L-PC']].max(axis=1,skipna=False)
        #df['ATR'] = df['TR'].rolling(n).mean()
        df['ATR'] = df['TR'].ewm(com=n,min_periods=n).mean()
        return df['ATR']
    
    def BOLLBND(DF,n=20):
        "function to calculate Bollinger Band"
        df = DF.copy()
        #df["MA"] = df['close'].rolling(n).mean()
        df["MA"] = df['Close'].ewm(span=n,min_periods=n).mean()
        df["BB_up"] = df["MA"] + 2*df['Close'].rolling(n).std(ddof=0) #ddof=0 is required since we want to take the standard deviation of the population and not sample
        df["BB_dn"] = df["MA"] - 2*df['Close'].rolling(n).std(ddof=0) #ddof=0 is required since we want to take the standard deviation of the population and not sample
        df["BB_width"] = df["BB_up"] - df["BB_dn"]
        df.dropna(inplace=True)
        return df
    
    def SettleOrders(Nthtime):
        unrealizedProfit = 0
        investedTotalAmount = 0
        LongOrShort = ""
        print("Checking if new Positions to settle...            DateTime:{}        Round: {} ".format(datetime.now(),Nthtime))
        dfPortfolioPositionsUtil = util.df(ib.reqPositions())
        if str(dfPortfolioPositionsUtil) == 'None':
            return
        dfPortfolioPositionsUtil = dfPortfolioPositionsUtil.drop(dfPortfolioPositionsUtil.loc[dfPortfolioPositionsUtil['position'] == 0].index)
        dfPortfolioPositionsUtil.drop_duplicates(inplace=True,ignore_index=True)
        i=0
        for stk in dfPortfolioPositionsUtil.values:
            PnLPercentage = 0
            PnL = 0
            symbol = dfPortfolioPositionsUtil["contract"].values[i].localSymbol
            if dfPortfolioPositionsUtil["account"].values[i] != accountIBKR:
                 i = i+1
                 continue
    
            alreadyOpenOrdersExists = False
            ib.reqAllOpenOrders()
            IsThereAnyPartialOpenOrders = False
            if dfPortfolioPositionsUtil["position"].values[i] > 0:
                LongOrShort = "Long "
                LongOrShortActionType = "BUY"
            else:
                LongOrShort = "Short"
                LongOrShortActionType = "SELL"
            dfOpenTrades = ib.openTrades()
    
            for trade in dfOpenTrades:
                if (trade.contract.localSymbol == symbol or trade.contract.symbol == symbol) and trade.order.action == LongOrShortActionType:
                    IsThereAnyPartialOpenOrders = True
                    break
               
           
            for trade in dfOpenTrades:
                 if (trade.contract.localSymbol == symbol or trade.contract.symbol == symbol) and trade.order.action != LongOrShortActionType:
                    if IsThereAnyPartialOpenOrders == True:
                        IsThereAnyPartialOpenOrders = True
                    else:
                        #ib.cancelOrder(trade.order)
                        alreadyOpenOrdersExists = True
                        break
                    #alreadyOpenOrdersExists = True
                    #break
            if alreadyOpenOrdersExists == True:
                print("An Open Order already exists for {}".format(symbol))
                i = i+1
                continue
    
            contract =Stock(symbol, 'SMART', 'USD')
            position = dfPortfolioPositionsUtil["position"].values[i]
            avgCost = dfPortfolioPositionsUtil["avgCost"].values[i]
            #avgCost = dfPortfolioPositionsUtil["avgCost"].values[i]
            mktDataPosition = ib.reqMktData(contract, "165,236,233,318", False, False, [])
            ib.sleep(1)
           
         
            if position != 0 :
                reTry = 0
                while (not (mktDataPosition.close > 0) and not reTry> maxTry):
                    mktDataPosition = ib.reqMktData(contract, "165,236,233,318", False, False, [])
                    ib.sleep(1)
                    reTry = reTry+1  
                if reTry >= maxTry:
                    i = i+1
                    continue
    
                decimalPlaces = len(str(mktDataPosition.close).split(".")[1])    
                lastPrice = mktDataPosition.close
                settleOrder = Order()
                settleOrder.account = dfPortfolioPositionsUtil["account"].values[i]
                if position > 0:
                     
                     lastPrice = mktDataPosition.bid
                     ProfitlimitPrice = str(round(avgCost + (profitTakerPercent/100)*avgCost, decimalPlaces))
                     investedTotalAmount = avgCost * position
                     unrealizedProfit = (profitTakerPercent/100)*avgCost * position
                     settleOrder.action = "SELL"
                     settleOrder.totalQuantity = position
                elif position < 0:
                   
                    lastPrice = mktDataPosition.ask
                    ProfitlimitPrice = str(round(avgCost - (profitTakerPercent/100)*avgCost, decimalPlaces))
                    investedTotalAmount = avgCost * position * -1
                    unrealizedProfit = (profitTakerPercent/100)*avgCost * position * -1
                    settleOrder.action = "BUY"
                    settleOrder.totalQuantity = position * -1
                settleOrder.account = dfPortfolioPositionsUtil["account"].values[i]
                settleOrder.orderType = "LMT"
                #settleOrder.lmtPrice = lastPrice
                settleOrder.lmtPrice = ProfitlimitPrice
                settleOrder.outsideRth = True
                settleOrder.usePriceMgmtAlgo = True
                settleOrder.tif = "GTC"
                #settleOrder.transmit = True
                if RunType == "Live":
                    dfsettleOrder = ib.placeOrder(contract,settleOrder)
                dictionary = {"Symbol": contract.symbol, "Position": position, "Price": avgCost}
                print("\n Placed Settle {} Order for Stock:{:6s} Qty:{:6s} Avg Price:{:8s} Amount:{:8s} Profit LMT Price:{:8s} Expected Profit%:{:3s} Profit Amt $:{:4s}".format(LongOrShort, symbol,str(int(position)),str(round(avgCost,decimalPlaces)),str(round(investedTotalAmount,2)),str(ProfitlimitPrice), str(profitTakerPercent), str(round(unrealizedProfit,2))))
            i=i+1
    
    
    if __name__ == "__main__":
        
        Nthtime = 1
        accountIBKR = "IBKRACCOUNTNUMBER" #str(sys.argv[1])
        ClientID = int("IBKRCLIENTNUMBER")
        # accountIBKR = str("U7329297")
        # ClientID = int(1)
        # print ("Im here ShortingOnTopGainersCMD.py")
        # print("accountIBKR", accountIBKR)
        # print("ClientID", ClientID)
        
        while True:
         try:
            os.system('cls' if os.name == 'nt' else 'clear')
            if str(ib) == '<IB not connected>':
                while str(ib) == '<IB not connected>':
                    try:
                        # print("Trying to Re-Connect.....")
                        ib.disconnect()
                        ib = IB()
                        #ib.connect('127.0.0.1', 7496, clientId=2) #paper trading port 7497 TWS Port
                        ib.connect('127.0.0.1', 4001, clientId=ClientID+40, timeout=60)
                        ib.sleep(1)
                    except Exception as ex:
                        pass
            
            result = ShortingOnTopGainers(Nthtime,accountIBKR,ClientID )
            # print(result)
            Nthtime = Nthtime + 1
            ib.sleep(1)
         except Exception as ex:
             ib.disconnect()          
             # print the error message
             print(f"An error occurred: {ex}")
             ib.sleep(1)   
             
         # finally:
             # ib.disconnect() 
             # print("Cleared Connections inside")
             # ib.sleep(30)  
except Exception as ex:
    ib.disconnect()          
    # print the error message
    print(f"An error occurred in ShortingOnTopGainersCMD.dat: {ex}")
    input("Press any key to exit...")
    ib.sleep(30)  
    
    
finally:
    ib.disconnect() 
    print("Clearing Connections and Clearing Memory: ShortingOnTopGainersCMD.dat")
    ib.sleep(30)     
    # sys.exit() # exit the script