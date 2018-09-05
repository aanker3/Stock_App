cd C:\Users\aanker\Desktop\Stock_App\Stock_Parser
set backupFilename=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%
echo %backupFilename%
copy StockList.xlsx StockSheetBackups\StockList_%backupFilename%_PreRun.xlsx
python StockParser.py
copy StockList.xlsx StockSheetBackups\StockList_%backupFilename%_PostRun.xlsx
pause