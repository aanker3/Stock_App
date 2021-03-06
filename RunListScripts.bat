set backupFilename=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%
echo %backupFilename%

cd C:\Users\aanker\Desktop\Stock_App\Stock_Parser
python StockParser.py
copy StockList.xlsx StockSheetBackups\StockList_%backupFilename%_PostRun.xlsx

cd C:\Users\aanker\Desktop\Stock_App\FinViz_Parser
python FinvizParser_20List.py Finviz_TLSupport_Oversold40_InvHammer
copy Finviz_TLSupport_Oversold40_InvHammer\Finviz_TLSupport_Oversold40_InvHammer.xlsx Finviz_TLSupport_Oversold40_InvHammer\StockSheetBackups\Finviz_TLSupport_Oversold40_InvHammer%backupFilename%_PostRun.xlsx

python FinvizParser_20List.py Finviz_WedgeStrng_Oversold40_AvgtrueRngUndr25
copy Finviz_WedgeStrng_Oversold40_AvgtrueRngUndr25\Finviz_WedgeStrng_Oversold40_AvgtrueRngUndr25.xlsx Finviz_WedgeStrng_Oversold40_AvgtrueRngUndr25\StockSheetBackups\Finviz_WedgeStrng_Oversold40_AvgtrueRngUndr25%backupFilename%_PostRun.xlsx

python FinvizParser_20List.py Finviz_MajorNews
copy Finviz_MajorNews\Finviz_MajorNews.xlsx Finviz_MajorNews\StockSheetBackups\Finviz_MajorNews%backupFilename%_PostRun.xlsx

python FinvizParser_20List.py Finviz_Upgrades
copy Finviz_Upgrades\Finviz_Upgrades.xlsx Finviz_Upgrades\StockSheetBackups\Finviz_Upgrades%backupFilename%_PostRun.xlsx

python FinvizParser_20List.py Finviz_TopGainers_PEless15
copy Finviz_TopGainers_PEless15\Finviz_TopGainers_PEless15.xlsx Finviz_TopGainers_PEless15\StockSheetBackups\Finviz_TopGainers_PEless15%backupFilename%_PostRun.xlsx

python FinvizParser_20List.py Finviz_Sma20Bounce
copy Finviz_Sma20Bounce\Finviz_Sma20Bounce.xlsx Finviz_Sma20Bounce\StockSheetBackups\Finviz_Sma20Bounce%backupFilename%_PostRun.xlsx

python FinvizParser_20List.py Finviz_OversoldEarningsMo
copy Finviz_OversoldEarningsMo\Finviz_OversoldEarningsMo.xlsx Finviz_OversoldEarningsMo\StockSheetBackups\Finviz_OversoldEarningsMo%backupFilename%_PostRun.xlsx

python FinvizParser_20List.py Finviz_Rsi30Reversal
copy Finviz_Rsi30Reversal\Finviz_Rsi30Reversal.xlsx Finviz_Rsi30Reversal\StockSheetBackups\Finviz_Rsi30Reversal%backupFilename%_PostRun.xlsx

python FinvizParser_20List.py Finviz_SmallCap_InstutionalTrans50Pct_30-45RSI_30SalesGrwth
copy Finviz_SmallCap_InstutionalTrans50Pct_30-45RSI_30SalesGrwth\Finviz_SmallCap_InstutionalTrans50Pct_30-45RSI_30SalesGrwth.xlsx Finviz_SmallCap_InstutionalTrans50Pct_30-45RSI_30SalesGrwth\StockSheetBackups\Finviz_SmallCap_InstutionalTrans50Pct_30-45RSI_30SalesGrwth%backupFilename%_PostRun.xlsx

pause