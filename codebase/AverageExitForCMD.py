from ib_insync import *
import pandas as pd
import xlwings as xw
import numpy as np
from IPython.display import display, clear_output
import asyncio
import nest_asyncio
nest_asyncio.apply()
import random
import ta
import sys
import time
import os
util.startLoop()
util.patchAsyncio()
ib = IB()

# ib.connect('127.0.0.1', 4001, clientId=int(9)) #paper trading port 7497 TWS Port
# ib.sleep(5)
# ib.disconnect()    
    
#accountIBKR= "U7504820"
accountIBKR= "U7329297"
#accountIBKR= "U6004156" 
AveragePercentFromLastPrice = 0.98

profitTakerPercent = 1
stopLossPercent = 70
RSIIndicatorAbove = 70.0
RSIDifferenceAbove = 40.0 
excel_file = r'C:\Vasanth\xlWingsAlgo\EncryptedFiles\Yazhi_Algo_Trading.xlsm'
sheet_name = 'Account Summary'
maxTry = 3


def RSI(DF,n=14):
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

#if 1==1:
#def AverageNowExit(accountIBKR,Ticker,PositionsExisting,PositionsToAverage, AverageExit_StatusXLCell,AverageExit_StatusXLCell ):
def AverageNowExit(accountIBKR,Ticker,PositionsExisting,PositionsToAverage,unrealized_PnL,  row ):
 ib = IB()
 try:
        
    AverageExit_StatusXLCell = "S" + row
    AverageExit_RSI_5Min_Back = "T" + row
    AverageExit_RSI_Now = "U" + row
    AverageExit_RSIDifference = "V" + row
    AverageTickerXLCell = "B" + row
    
    
    
   
    
    if ib.isConnected():
        ib.disconnect()
    ib.connect('127.0.0.1', 4001, clientId=int(row)+12) #paper trading port 7497 TWS Port
    ib.sleep(5)
    contract=Stock(Ticker, 'SMART', 'USD')
    
    # if str(ib) == '<IB not connected>':
    #     while str(ib) == '<IB not connected>':
    #         try:
    #             print("Trying to Re-Connect.....")
    #             ib.disconnect()
    #             ib = IB()
    #             #ibClient1.connect('127.0.0.1', 7496, clientId=2) #paper trading port 7497 TWS Port
    #             ib.connect('127.0.0.1', 4001, clientId=int(row+1))
    #         except Exception as ex:
    #             ib.disconnect()
    #             pass
    
    # mktData = ib.reqMktData(contract, "236", False, False, []) 
    # ib.sleep(1)
    # last_price = round(mktData.last if not np.isnan(mktData.last) else mktData.close,deci)

    PositionsTotalToBe = PositionsToAverage + PositionsExisting
    ib.reqAllOpenOrders()
    for trade in ib.openTrades():
        #print(trade.contract.symbol)
        if (trade.contract.localSymbol == Ticker or trade.contract.symbol == Ticker):
            #if trade.order.lmtPrice > avgCost:
                ib.cancelOrder(trade.order)
                print("Cancelled all orders for: ", Ticker) 

    wb = xw.Book(excel_file)
    sheet = wb.sheets[sheet_name]
    IsAverageSettleOrderPlace = False
    
    #print (AverageTickerXLCell)
    while not IsAverageSettleOrderPlace:
         if sheet.range(AverageTickerXLCell).value != Ticker:
             IsAverageSettleOrderPlace = True
             break
         dfPortfolioPositionsUtil = util.df(ib.reqPositions())
         
         if str(dfPortfolioPositionsUtil) != 'None':
            dfPortfolioPositionsUtil = dfPortfolioPositionsUtil.drop(dfPortfolioPositionsUtil.loc[dfPortfolioPositionsUtil['position'] == 0].index)
            i=0
            
            os.system('cls' if os.name == 'nt' else 'clear')
            
            #print("checking Portfolio")
            for Portstock in dfPortfolioPositionsUtil.values:
                if Portstock[1].symbol == Ticker: # or Portstock[0] != accountIBKR:
                    print("Yazhi Algo Trading trying to average & exit:",  Ticker)
                   
                    positionNow = dfPortfolioPositionsUtil["position"].values[i] 
                    avgCost = dfPortfolioPositionsUtil["avgCost"].values[i]
                  
                    deci = 2
                    if avgCost >= 1:
                        deci = 2
                    else:
                        deci = 4
                    avgCost = round (avgCost,deci)   
                    
                    base_price = round(positionNow * avgCost, deci)
                    print("IBKR Account: ", accountIBKR, "Ticker:", Ticker)
                    ib.sleep(1)
                    print("         Positions          :" , positionNow)
                    ib.sleep(0.5)
                    print("         Average Cost       :" , avgCost)
                    ib.sleep(0.5)
                    print("         Base Price         :" , base_price)
                    ib.sleep(0.5)
                    #print("Checking Current Position", Ticker, positionNow)
                    
                    if PositionsTotalToBe >= positionNow: #and float(Unrealized_PnL) <= 0 :
                        sheet.range(AverageExit_StatusXLCell).value  =   "Trying to Average"
                        sheet.range(AverageExit_StatusXLCell).color = "#8fce00"
                        ib.sleep(1)
                        sheet.range(AverageExit_StatusXLCell).color = None
                        ConditionSatisfied = 'No'
                        
                        # trying to average the Condtion Satisfy
                        print("Getting Market data for", Ticker)
                        mktData = ib.reqMktData(contract, "165,236,233,318", False, False, [])
                        ib.sleep(1)
                        halted = mktData.halted #Volatility halt = 2
                        

                        reTry = 0
                        while (not (mktData.last >0  and mktData.close > 0 and mktData.bid > 0 and mktData.ask > 0) and not reTry> maxTry):
                            mktData = ib.reqMktData(contract, "165,236,233,318", False, False, [])
                            print("Unable to get Market data. Retry: ", reTry)
                            ib.sleep(1)
                            
                            reTry = reTry+1
                        if reTry >= maxTry:
                            i = i+1
                            continue
                                  
                        prvClosePrice = mktData.close
                        lastPrice = mktData.last
                        bidPrice = mktData.bid
                        askPrice = mktData.ask
                        highPrice = mktData.high
                        lowPrice = mktData.low
                        
                        market_value = round(positionNow* lastPrice, deci)
                        Unrealized_PnL = round(market_value - base_price ,2) #((lastPrice/ position.avgCost) - 1) * 100 * -1
                        Unrealized_PnL_perc = round((Unrealized_PnL / base_price) * 100 *-1,2)
                        new_avg_price = lastPrice * AveragePercentFromLastPrice
                        # #PositionsToAverage =(Base Price - (New Average Price * Positions))/(New Average Price - Last Price)
                        # PositionsToAverage =  (base_price - (new_avg_price * positionNow))/(new_avg_price-lastPrice)
                        # How_Much_More_Amount = PositionsToAverage *  lastPrice 
                        # PositionsTotalToBe = positionNow + PositionsToAverage 
                        # Total_Base_Price = base_price + How_Much_More_Amount
                        # Total_Market_Price = Total_Positions * lastPrice
                        
                        # print("IBKR Account: ", accountIBKR, "Ticker:", Ticker)
                        # print("         Positions          :" , positionNow)
                        print("         Positions to Avg   :" , PositionsToAverage)
                        print("         last Price         :" , lastPrice)
                        print("         Base Price         :" , base_price)
                        print("         Market Value       :" , market_value)
                        print("         UnRealized PnL     :" , Unrealized_PnL)
                        print("         UnRealized PnL Perc:" , Unrealized_PnL_perc)
                        ib.sleep(1)
                        #print (Unrealized_PnL)
                        if Unrealized_PnL >= 0:
                           i = i+1
                           continue 
                        #alreadyOpenOrdersExists = False
                        if halted == 2 and lastPrice > Price:
                                alreadyOpenOrdersExists = False
                                ib.reqAllOpenOrders()
                                for trade in ib.openTrades():
                                    #print(trade.contract.symbol)
                                    if (trade.contract.localSymbol == Ticker or trade.contract.symbol == Ticker) and trade.order.action == "SELL" and trade.order.orderType == "MKT":
                                        alreadyOpenOrdersExists = True
                                        break 
                                        
                                if alreadyOpenOrdersExists == True:
                                    print("An Open Order already exists for {}".format(Stocksymbol))
                                    i = i+1
                                    continue            
                                if alreadyOpenOrdersExists == False:      
                                    print("     ******Condition Satisfied******")
                                    ConditionSatisfied = 'Yes'
                                    shortOrder = Order()
                                    shortOrder.account = accountIBKR
                                    shortOrder.action = "SELL"
                                    shortOrder.orderType = "MKT"
                                    shortOrder.totalQuantity = PositionsToAverage
                                    #shortOrder.lmtPrice = limitPrice
                                    #shortOrder.outsideRth = True
                                    shortOrder.usePriceMgmtAlgo = True
                                    shortOrder.transmit = True
                                    dfshortOrder = ib.placeOrder(contract,shortOrder)
                                    print("     ******Condition Satisfied****** Halted *** MKT Order")
                                    i = i+1
                                    continue 
                        #histData = ibClient1.reqHistoricalData(contract,'','900 S','30 secs','ADJUSTED_LAST',1,1,0,[])      
                        # Request historical data for last 10 minutes at 5-minute interval
                        bars = ib.reqHistoricalData(
                            contract,
                            endDateTime='',
                            durationStr='600 S',
                            barSizeSetting='30 secs',
                            whatToShow='TRADES',
                            useRTH=False,
                            formatDate=1
                            )   
                        ib.sleep(0.5)
                        #print (bars)
                        # # Convert the returned bars to a DataFrame
                        # df_bars = util.df(bars)
                        # df_bars.set_index('date', inplace=True)
                        # # Calculate RSI for the 5th minute
                        # RSIndiacator5MinBack = ta.RSI(df_bars[:5].close, timeperiod=14)
                        # # Calculate RSI for the last minute
                        # RSIIndicatorLatest = ta.RSI(df_bars[-1:].close, timeperiod=14)
                        # RSIDifference = RSIIndicatorLatest - RSIndiacator5MinBack
                        dfUtilhisData = util.df(bars)
                        
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
                          
                        RSIIndicatorLatest =  round(dfRSI.values[len(dfRSI)-1],0)
                        RSIDifference = RSIIndicatorLatest - RSIndiacator5MinBack
                        
                        print("         RSI 5 Min back     :" , RSIndiacator5MinBack)
                        print("         RSI Now            :" , RSIIndicatorLatest)
                        print("         RSI Difference     :" , RSIDifference)
                        ib.sleep(0.5)
                        
                        sheet.range(AverageExit_RSI_5Min_Back).value  =   RSIndiacator5MinBack
                        sheet.range(AverageExit_RSI_5Min_Back).color = "#8fce00"
                        ib.sleep(0.5)
                        sheet.range(AverageExit_RSI_5Min_Back).color = None
                        
                        sheet.range(AverageExit_RSI_Now).value  =   RSIIndicatorLatest
                        sheet.range(AverageExit_RSI_Now).color = "#8fce00"
                        ib.sleep(0.5)
                        sheet.range(AverageExit_RSI_Now).color = None
                        
                        
                        sheet.range(AverageExit_RSIDifference).value  =   RSIDifference
                        sheet.range(AverageExit_RSIDifference).color = "#8fce00"
                        ib.sleep(0.5)
                        sheet.range(AverageExit_RSIDifference).color = None
                        
                        
                        if  ((RSIIndicatorLatest >  RSIndiacator5MinBack and RSIIndicatorLatest >=  RSIIndicatorAbove ) or (RSIDifference > RSIDifferenceAbove) ):
                               print("     ******Condition Satisfied******")
                               ConditionSatisfied = 'Yes'
                               
                               ib.reqAllOpenOrders()
                               for trade in ib.openTrades():
                                   #print(trade.contract.symbol)
                                   if (trade.contract.localSymbol == Ticker or trade.contract.symbol == Ticker):
                                       #if trade.order.lmtPrice > avgCost:
                                           ib.cancelOrder(trade.order)
                                           #print('Cancelled all orders for: ", Ticker) 
                               
                               AverageOrder = Order()
                               AverageOrder.account = accountIBKR
                               AverageOrder.action = "SELL"
                               AverageOrder.orderType = "LMT"
                               AverageOrder.totalQuantity = PositionsToAverage
                               AverageOrder.lmtPrice = lastPrice
                               AverageOrder.outsideRth = True
                               AverageOrder.usePriceMgmtAlgo = True
                               AverageOrder.transmit = True
                               dfAverageOrder= ib.placeOrder(contract,AverageOrder)
                               ib.sleep(1)
                
                    

                    elif PositionsTotalToBe <= positionNow and Unrealized_PnL > 0:
                        alreadyOpenOrdersExists = False
                        ib.reqAllOpenOrders()
                        for trade in ib.openTrades():
                            #print(trade.contract.symbol)
                            if (trade.contract.localSymbol == Ticker or trade.contract.symbol == Ticker) and trade.order.action == "BUY":
                                #if trade.order.lmtPrice > avgCost:
                                    ib.cancelOrder(trade.order)
                                    print('Cancelled existing order for exit') 

                        mktData = ib.reqMktData(contract, "165,236,233,318", False, False, [])
                        #deci = len(str(mktData.close).split(".")[1])   
                        #deci = 2
                        ib.sleep(1)
                        
                        settleOrder = Order()
                        settleOrder.account = dfPortfolioPositionsUtil["account"].values[i]
                        investedTotalAmount = avgCost * positionNow * -1
                        unrealizedProfit = (profitTakerPercent/100)*avgCost * positionNow * -1
                        ProfitlimitPrice = str(round(avgCost - (profitTakerPercent/100)*avgCost, deci))
                        settleOrder.action = "BUY"
                        settleOrder.totalQuantity = positionNow 
                        settleOrder.orderType = "LMT"
                        settleOrder.lmtPrice = str(round(avgCost - (profitTakerPercent/100)*avgCost, deci))
                        settleOrder.outsideRth = True
                        settleOrder.usePriceMgmtAlgo = True
                        settleOrder.tif = "GTC"
                        settleOrder.transmit = True
                        dfsettleOrder = ib.placeOrder(contract,settleOrder)
                        IsAverageSettleOrderPlace = True
                        
                        sheet.range(AverageExit_StatusXLCell).value  =   "Exit Order PLaced"
                        sheet.range(AverageExit_StatusXLCell).color = "#8fce00"
                        ib.sleep(1)
                        sheet.range(AverageExit_StatusXLCell).color = None
     
     
     
                        print("\n Placed Settle {} Order for Stock:{:6s} Qty:{:6s} Avg Price:{:8s} Amount:{:8s} Profit LMT Price:{:8s} Expected Profit%:{:3s} Profit Amt $:{:4s}".format('Short', Ticker,str(int(positionNow)),str(round(avgCost,deci)),str(round(investedTotalAmount,2)),str(ProfitlimitPrice), str(profitTakerPercent), str(round(unrealizedProfit,2))))
                        
                i = i+1

    ib.disconnect()                
    print("Completed Average & Exit")
    ib.sleep(30)        
 except Exception as ex:
    ib.disconnect()          
    # print the error message
    print(f"An error occurred: {ex}")
    ib.sleep(30)   
    
 finally:
    ib.disconnect() 
    ib.sleep(30)   
# # #AverageNowExit(accountIBKR,Ticker,PositionsExisting,PositionsToAverage,Unrealized_PnL,  row ):
if __name__ == "__main__":
    accountIBKR = str(sys.argv[1])
    Ticker = str(sys.argv[2])
    PositionsExisting = int(sys.argv[3])
    PositionsToAverage = int(sys.argv[4])
    Unrealized_PnL = int(sys.argv[5])
    row = str(sys.argv[6])
    time.sleep(10)
    result = AverageNowExit(accountIBKR,Ticker,PositionsExisting,PositionsToAverage,Unrealized_PnL,  row )
    print(result)
    time.sleep(60)
# AverageNowExit(accountIBKR,Ticker,PositionsExisting,PositionsToAverage,unrealized_PnL,  row )
#AverageNowExit(str("U7329297"), str("OCFT"), -152, 852, -106,  str(8))