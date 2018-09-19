set backupFilename=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%
echo %backupFilename%

cd C:\Users\aanker\Desktop\Stock_App\FinViz_Parser

python FinvizParser.py Finviz_Upgrades
copy Finviz_Upgrades\Finviz_Upgrades.xlsx Finviz_Upgrades\StockSheetBackups\Finviz_Upgrades%backupFilename%_PostRun.xlsx

pause