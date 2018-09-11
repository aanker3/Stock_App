set backupFilename=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%
echo %backupFilename%

cd C:\Users\aanker\Desktop\Stock_App\FinViz_Parser\Finvis_DoubleBot
python Finviz_DoubleBot.py
copy Finviz_DoubleBot.xlsx StockSheetBackups\Finviz_DoubleBot_%backupFilename%_PostRun.xlsx

cd C:\Users\aanker\Desktop\Stock_App\FinViz_Parser\Finvis_StrongBuy
python Finviz_StrongBuy.py
copy Finviz_StrongBuy.xlsx StockSheetBackups\Finviz_StrongBuy_%backupFilename%_PostRun.xlsx

cd C:\Users\aanker\Desktop\Stock_App\Stock_Parser
python StockParser.py
copy StockList.xlsx StockSheetBackups\StockList_%backupFilename%_PostRun.xlsx

pause