set backupFilename=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%
echo %backupFilename%

cd C:\Users\aanker\Desktop\Stock_App\FinViz_Parser\Finviz_DoubleBot
python Finviz_DoubleBot.py
copy Finviz_DoubleBot.xlsx StockSheetBackups\Finviz_DoubleBot_%backupFilename%_PostRun.xlsx

cd C:\Users\aanker\Desktop\Stock_App\FinViz_Parser\Finviz_TLSupport
python Finviz_TLSupport.py
copy Finviz_TLSupport.xlsx StockSheetBackups\Finviz_TLSupport_%backupFilename%_PostRun.xlsx

cd C:\Users\aanker\Desktop\Stock_App\FinViz_Parser\Finviz_InsiderBuying_BuyPlus
python Finviz_InsiderBuying_BuyPlus.py
copy Finviz_InsiderBuying_BuyPlus.xlsx StockSheetBackups\Finviz_InsiderBuying_BuyPlus_%backupFilename%_PostRun.xlsx

cd C:\Users\aanker\Desktop\Stock_App\FinViz_Parser\Finviz_ChannelUp
python Finviz_ChannelUp.py
copy Finviz_ChannelUp.xlsx StockSheetBackups\Finviz_ChannelUp_%backupFilename%_PostRun.xlsx

cd C:\Users\aanker\Desktop\Stock_App\Stock_Parser
python StockParser.py
copy StockList.xlsx StockSheetBackups\StockList_%backupFilename%_PostRun.xlsx

pause