#Shorting on Top gainers Everday 
#**Program Starts here**

import xlwings as xw
from ib_insync import *
import pandas as pd
import os
from datetime import datetime, timedelta, time
import pytz
import subprocess
import math
import sys
import tempfile
from Crypto.Cipher import AES
util.startLoop()

# Write to Excel file
excel_file = r'{Yaali_Algo_Trading_Full_FilePath}'
# excel_file = 'C:\\Vasanth\\xlWingsAlgo\\EncryptedFiles\\Yazhi_Algo_Trading.xlsm'
sheet_name = '{Yaali_Algo_Sheet_Name}'
# excel_file = r"[{ExcelFileLocation}]"
# sheet_name = "[{ExcelSheetName}]"

wb = xw.Book(excel_file)
sheet = wb.sheets[sheet_name]

rootPath = sheet.range('M2').value
Puthon_Path = sheet.range('M3').value
# print(rootPath)
data_folder = f'{rootPath}/datafiles'
codebase_folder =  f'{rootPath}/codebase'
log_folder =  f'{rootPath}/logs'
archive_folder = f'{data_folder}/archive'
FeeRate_folder = f'{data_folder}/FeeRate'
today = datetime.now().strftime('%Y_%m_%d')

ShortingOnTopGainersMonitor = "ShortingOnTopGainersCMD.py"
IBKRFeeRateFTP = "USStockList_IBKR_FeeRate_ShortableShares_ftp_v2.dat"
Decryption_Key =  "c2e5b8a3d14ad9f43ab53a4c2d3ed4f5"

def get_market_status(exchange_timezone: str = 'America/New_York'):
    """
    Check if the market is in pre-market, open, or after-hours status.

    :param exchange_timezone: Timezone of the market. Default is 'America/New_York'.
    :return: Market status as a string.
    """
    # Define market hours (example for NYSE)
    market_open = time(9, 30)  # 9:30 AM
    market_close = time(16, 0)  # 4:00 PM
    pre_market_start = time(4, 0)  # 4:00 AM
    after_hours_end = time(20, 0)  # 8:00 PM

    # Get the current time in the specified timezone
    timezone = pytz.timezone(exchange_timezone)
    now = datetime.now(timezone)  # Localized current time
    
    if pre_market_start <= now.time() < market_open:
        return "Pre-Market"
    elif market_open <= now.time() <= market_close:
        return "Market Open"
    elif market_close < now.time() <= after_hours_end:
        return "After-Hours"
    else:
        return "Market Closed"
    
def DecryptFile(Encrypted_File, Decryption_Key, ):
    decryption_key = str(Decryption_Key).encode()
    Encrypted_File_Path = os.path.join(codebase_folder, Encrypted_File) 
    try:
        with open(Encrypted_File_Path, 'rb') as f:
            encrypted_contents = f.read()
        cipher = AES.new(decryption_key, AES.MODE_CBC)
        decrypted_contents = cipher.decrypt(encrypted_contents)
        padding_len = decrypted_contents[-1]
        decrypted_contents = decrypted_contents[:-padding_len]
        decrypted_contents = decrypted_contents.split(b'\n', 2)[2]
        # exec(decrypted_contents, globals())
        return decrypted_contents
    except Exception as e:
        return ""

def get_contents(filename, codebase_folder):
    """
    Reads and returns the content of a Python file 
    located inside the given codebase folder.

    :param filename: str - Python file name (e.g. "ShortingOnTopGainersCMD.py")
    :param codebase_folder: str - Path to the codebase directory
    :return: str - Content of the file
    """
    file_path = os.path.join(codebase_folder, filename)
    
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read()
    
def get_first_empty_row_in_column_b():
    # Open the Excel file and select the sheet
    wb = xw.Book(excel_file)
    sheet = wb.sheets[sheet_name]

    # Loop through B9 to B109 to find the first empty cell
    for row in range(9, 110):  # from row 9 to 109 (inclusive)
        cell_value = sheet.range(f'B{row}').value
        if cell_value is None or cell_value == "":
            return row  # Return the first empty row number
            
# Connect to TWS
ib = IB()

decrypted_contents = ""
# script_path = os.path.realpath(sys.argv[0])
# print("Executing: ", script_path)
monitoringCount = 0
try:

    # decrypted_contents = DecryptFile(IBKRFeeRateFTP,Decryption_Key)
    # exec(decrypted_contents, globals())
    
    Monitor_Contents = get_contents(ShortingOnTopGainersMonitor,codebase_folder)
    # Monitor_Contents = Monitor_Contents.decode()
    Monitor_Contents = Monitor_Contents.replace("[Yaali_Algo_Trading_Full_FilePath]", excel_file)
    Monitor_Contents = Monitor_Contents.replace("[Yaali_Algo_Sheet_Name]", sheet_name)
    # exec(ShortingOnTopGainersMonitor_Contents, globals())

    IsFirstTime = True
    if (sheet.range('J7').value == "Exit"):
        sheet.range('J7').value    = "Show in Excel"
    accountIBKR= sheet.range(f'C2').value 
    show_type = sheet.range('J7').value    
    print("Shorting On Top Gainers for IBKR Account: ", accountIBKR)
    print("Running in...", show_type)
    ib.sleep(0.5)
    ib.connect('127.0.0.1', 4001, clientId=40, timeout=60)
    # ib.connect('127.0.0.1', 4001, clientId=40)
    print("IBKR Gateway Client Connected")
    
    while (sheet.range('J7').value != "Exit"):
        
        if IsFirstTime == False:
            show_type_changed = sheet.range('J7').value 
            sheet.range('J7').value  = "Exit"  
            ib.sleep(2)
            sheet.range('J7').value = show_type_changed
            # print("Outer Loop" , sheet.range('J7').value)  
        start_row = 8
        # sheet.range(f'A9:Z31').clear_contents()
        sheet.range(f'A9:Z31').value = None
        
        sheet.range(f'B9:Z31').color = None
        show_type = sheet.range('J7').value    
        
        dfUniverse = pd.DataFrame(columns=["S.No", "Ticker", "Volume", "Prev Close", "Last", "ClientID"])
        client_id = 0
        while (show_type == sheet.range('J7').value) or sheet.range('J7').value == "Exit":
            # print("Inner Loop" , sheet.range('J7').value)  
            show_type = sheet.range('J7').value    
            monitoringCount = monitoringCount+1
            
            if show_type == 'Exit':
                print ("Show Type choosen as Exit. If you want to keep run Change the Method of Show type")
                sheet.range(f'B9:W24').color = None
                ib.disconnect()
                sys.exit()

            MaxPositionsInPortfolio = sheet.range(f'K2').value 
            maxCashPerStock = sheet.range(f'K3').value 
            profitTakerPercent =  sheet.range(f'K4').value 
            stopLossPercent =  sheet.range(f'K5').value 
            HowManyTickersForEachClient =  sheet.range(f'K6').value 
            
            entryPercentage = sheet.range(f'E3').value 
            priceBelow =  sheet.range(f'E6').value 
            priceAbove =  sheet.range(f'E7').value 
            
            market_status = get_market_status()
            sheet.range(f'M4').value = market_status
            if market_status == "Pre-Market":
                 avgVolumeAbove =  1000 
                 volumeAbove = 100000
            else:
                # avgVolumeAbove =  sheet.range(f'E4').value 
                # volumeAbove = sheet.range(f'E5').value 
                avgVolumeAbove =  sheet.range(f'E4').value 
                volumeAbove = 1000000
                
          
            
            #print(priceAbove)
            RSIAbove =  sheet.range(f'G2').value 
            RSIAboveifLastPriceLessOneDollar =sheet.range(f'G3').value    
            RSIDifferenceAbove = sheet.range(f'G5').value          
            #print("here")
            subsTopGainers = ScannerSubscription(instrument = "STK", locationCode = "STK.US.MAJOR", scanCode = "TOP_PERC_GAIN" )
            tagvalues = []
            
            
            
        
            # tagvalues.append(TagValue("avgVolumeAbove", avgVolumeAbove));
            tagvalues.append(TagValue("priceBelow", priceBelow));
            tagvalues.append(TagValue("priceAbove", priceAbove));
            tagvalues.append(TagValue("changePercAbove", entryPercentage));
            tagvalues.append(TagValue("volumeAbove", volumeAbove));
            # print("market_status", market_status)
            # print("avgVolumeAbove", avgVolumeAbove)
            # print("volumeAbove", volumeAbove)
            #print("scannerData")
            scannerData = ib.reqScannerData(subsTopGainers,[],tagvalues)
            ib.sleep(1)
            #print(scannerData)
            # ib.sleep(5)
            os.system('cls' if os.name == 'nt' else 'clear')    
            print("Shorting On Top Gainers for IBKR Account: ", accountIBKR)
            print("Running in...", show_type, "Monitoring Count:", monitoringCount)
            tickers = [scan.contractDetails.contract.symbol for scan in scannerData]

            print("Scanned Tickers:", tickers)    
            
            ManualTickersSet_1 = sheet.range('N2:N7').value
            ManualTickersSet_2 = sheet.range('O2:O7').value
            ManualTickersSet_3  = sheet.range('P2:P7').value
            
            ManualTickers = []
            if ManualTickersSet_1:
                ManualTickers.extend(ManualTickersSet_1)
            if ManualTickersSet_2:
                ManualTickers.extend(ManualTickersSet_2)
            if ManualTickersSet_3:
                ManualTickers.extend(ManualTickersSet_3)
            
            ManualTickers = [value for value in ManualTickers if value is not None]
            
            
            for value in ManualTickers:
                if value not in tickers:
                    tickers.append(value)
            
            ExcludeTickersSet_1 = sheet.range('Q2:Q7').value
            ExcludeTickersSet_2 = sheet.range('R2:R7').value
            ExcludeTickersSet_3  = sheet.range('S2:S7').value
           
                    
            ExcludeTickers = []
            if ExcludeTickersSet_1:
                ExcludeTickers.extend(ExcludeTickersSet_1)
            if ExcludeTickersSet_2:
                ExcludeTickers.extend(ExcludeTickersSet_2)
            if ExcludeTickersSet_3:
                ExcludeTickers.extend(ExcludeTickersSet_3)
            
            ExcludeTickers = [value for value in ExcludeTickers if value is not None]
            

            print("ExcludeTickers:", ExcludeTickers)        
            for value in ExcludeTickers:
                if value in tickers:
                    tickers.remove(value)
            tickers = [ticker for ticker in tickers if len(ticker) != 5 and "." not in ticker and " " not in ticker]

            #print(tickers)
            for i, ticker in enumerate(tickers):
                excel_tickers = [sheet.range(f'B{row}').value for row in range(9, 110) if sheet.range(f'B{row}').value]
                if ticker not in excel_tickers: 
                    empty_row = get_first_empty_row_in_column_b()
                    sheet.range(f'A{empty_row}').value = empty_row - 8
                    sheet.range(f'B{empty_row}').value = ticker
                    sheet.range(f'C{empty_row}').value = "Monitoring"
                    ib.sleep(1)  
                   
                    if i % HowManyTickersForEachClient == 0:
                        client_id = empty_row - 8
                        ClientContent = Monitor_Contents
                        ClientContent = ClientContent.replace("IBKRACCOUNTNUMBER", accountIBKR)
                        ClientContent = ClientContent.replace("IBKRCLIENTNUMBER", str(client_id))
                        
                        print("Creating new Client to Monitor ", ticker)
                        with tempfile.NamedTemporaryFile(mode='w', delete=False) as f:
                            f.write(ClientContent)
                            tmp_file_path = f.name
                        
                     
                        if show_type != 'Show in Console':
                            subprocess.Popen(['python', tmp_file_path], stdout=open(os.devnull, 'w'), stderr=open(os.devnull, 'w'))
                        else:
                            subprocess.Popen(['start', 'cmd', '/c',Puthon_Path, tmp_file_path], shell=True)
                        ib.sleep(2)   
                        

except Exception as ex:
    # print the error message
    print(f"Error in ShortingOnTopGainersList.dat: {ex}")
    # ib.sleep(1)  
    ib.disconnect() 
    # input("Press any key to exit...")
    sys.exit()
    
    
finally:
    print("Clearing Connections and Clearing Memory: ShortingOnTopGainersList.dat")
    ib.disconnect()
    sheet.range('J7').value  = "Exit"  
    # ib.sleep(5)   
    sys.exit() # exit the script



# import subprocess
# temp_py_file_path = r"C:\\Users\\thang\\AppData\\Local\\Temp\\tmprzomuf3_"
# subprocess.Popen(['start', 'cmd', '/c', 'python', temp_py_file_path], shell=True)
# subprocess.Popen(["C:\\ProgramData\\anaconda3\\python.exe", temp_py_file_path])



# # fileName = r"C:\Windows\Temp\tmpc6_gtsqi.py"


# subprocess.Popen(['start', 'cmd', '/c', 'python', fileName], shell=True)
