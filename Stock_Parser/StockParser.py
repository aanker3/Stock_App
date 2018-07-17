#!/usr/bin/env python3

# Name            :  StockParser
# Script Name     :  StockParser.py

from datetime import datetime
from lxml import html
import requests
import time
import Tkinter
from PIL import ImageTk, Image
import os.path
import subprocess


def StockParse():
        # Get Stock info
        page = requests.get('https://eresearch.fidelity.com/eresearch/gotoBL/fidelityTopOrders.jhtml', verify=False)
        tree = html.fromstring(page.content)
 
        sTicker= ["1"] * 30
        sPrice = ["1"] * 30
        sBuy = ["1"] * 30
        sSell = ["1"] * 30
        sBuySellRatio = ["1"] * 30        

        for i in range (1,31):
            TickerListString = "//*[@id=\"topOrdersTable\"]/tbody/tr["
            TickerListString += str(i)
            TickerListString += "]/td[2]/span"

            PriceListString = "//*[@id=\"topOrdersTable\"]/tbody/tr["
            PriceListString += str(i)
            PriceListString += "]/td[4]/span"

            buyOrderListString = "//*[@id=\"topOrdersTable\"]/tbody/tr["
            buyOrderListString += str(i)
            buyOrderListString += "]/td[5]/span"
            
            sellOrderListString = "//*[@id=\"topOrdersTable\"]/tbody/tr["
            sellOrderListString += str(i)
            sellOrderListString += "]/td[7]/span"            

            sTicker[i-1] = tree.xpath(TickerListString)[0].text
            sPrice[i-1] = tree.xpath(PriceListString)[0].text
            sBuy[i-1] = tree.xpath(buyOrderListString)[0].text
            sSell[i-1] = tree.xpath(sellOrderListString)[0].text
            sBuySellRatio[i-1] = round((float(sBuy[i-1]) / (float(sBuy[i-1]) + float(sSell[i-1]))),2)
        print 'sTicker ', sTicker
        print 'sPrice ', sPrice
        print 'sBuy ', sBuy
        print 'sSell ', sSell
        print 'sBuySellRatio ', sBuySellRatio        
        
#        stock_Ticker = tree.xpath('//*[@id="topOrdersTable"]/tbody/tr[3]/td[2]/span')[0].text
#        Stocks = tree.xpath('//*[@id="topOrdersTable"]/tbody/tr[3]/td[3]')[0].text
#        print 'Stock Ticker: ', stock_Ticker
#        print 'Stocks: ', Stocks
#        print("Done")



def main():
    
    return StockParse()



if __name__ == '__main__':
    main()