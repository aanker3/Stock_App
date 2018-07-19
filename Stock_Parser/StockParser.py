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

#google and morning star are not working for some reason.  List out of range.  has to do with the [0].text.  Not sure

#def GetGoogle_MainPage(ticker):
#https://www.google.com/search?tbm=fin&ei=2ntPW7LGGM-p_QbjpqyYAQ&q=goog&oq=goog&gs_l=finance-immersive.3..81l3.9658.10169.0.10352.4.4.0.0.0.0.225.225.2-1.1.0..1..0...1.1.64.finance-immersive..3.1.224....0.lQqDD8e_eWs#scso=uid_5XtPW6ifKeSkgge-6oc4_5:0,uid_9XtPW-7SC8PM_AaT5qawAg_5:0
#    pageStr="https://www.google.com/search?tbm=fin&ei=2ntPW7LGGM-p_QbjpqyYAQ&q="+ticker+"&oq="+ticker+"&gs_l=finance-immersive.3..81l3.9658.10169.0.10352.4.4.0.0.0.0.225.225.2-1.1.0..1..0...1.1.64.finance-immersive..3.1.224....0.lQqDD8e_eWs#scso=uid_5XtPW6ifKeSkgge-6oc4_5:0,uid_9XtPW-7SC8PM_AaT5qawAg_5:0"
#    page = requests.get(pageStr)
#    tree = html.fromstring(page.content)
#    return tree
    
#def GetMorningStar_MainPage(ticker):
#https://www.google.com/search?tbm=fin&ei=2ntPW7LGGM-p_QbjpqyYAQ&q=goog&oq=goog&gs_l=finance-immersive.3..81l3.9658.10169.0.10352.4.4.0.0.0.0.225.225.2-1.1.0..1..0...1.1.64.finance-immersive..3.1.224....0.lQqDD8e_eWs#scso=uid_5XtPW6ifKeSkgge-6oc4_5:0,uid_9XtPW-7SC8PM_AaT5qawAg_5:0
#    pageStr="https://www.google.com/search?tbm=fin&ei=2ntPW7LGGM-p_QbjpqyYAQ&q="+ticker+"&oq="+ticker+"&gs_l=finance-immersive.3..81l3.9658.10169.0.10352.4.4.0.0.0.0.225.225.2-1.1.0..1..0...1.1.64.finance-immersive..3.1.224....0.lQqDD8e_eWs#scso=uid_5XtPW6ifKeSkgge-6oc4_5:0,uid_9XtPW-7SC8PM_AaT5qawAg_5:0"
#    page = requests.get(pageStr)
#    tree = html.fromstring(page.content)
#    return tree
    
#def GetGoogleStock_Price(ticker):
    #tree = GetGoogle_MainPage(ticker)
    #price = round(float((tree.xpath('//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[1]/span[1]/span/span[1]')[0].text).replace(",","")),2)
    #return price

#def GetGoogleStock_PriceChange(ticker):
    #tree = GetGoogle_MainPage(ticker)
    #priceChange = tree.xpath('//*[@id="knowledge-finance-wholepage__entity-summary"]/div/g-card-section/div/g-card-section/div[1]/span[2]/span[1]')[0].text
    #return priceChange    
    
def GetYahoo_MainPage(ticker):
    #https://finance.yahoo.com/quote/AAPL?p=AAPL&.tsrc=fin-srch-v1
    pageStr="https://finance.yahoo.com/quote/"+ticker+"?p="+"ticker+&.tsrc=fin-srch-v1"
    page = requests.get(pageStr)
    tree = html.fromstring(page.content)
    return tree
    
def GetYahooStock_Price(tree):
    price = round(float((tree.xpath('//*[@id="quote-header-info"]/div[3]/div[1]/div/span[1]')[0].text).replace(",","")),2)
  #  print 'price : ', price
    return price
    
def GetYahooStock_PriceChange(tree):
    #note: gives data in format of -.2717 (-.01%)  This funciton grabs would grab and return -.27 in this situation
    priceChange = tree.xpath('//*[@id="quote-header-info"]/div[3]/div[1]/div/span[2]')[0].text
    priceChange = round(float(priceChange.partition(' ')[0]),2)


   
    return priceChange   
    
def GetYahooStock_Beta(tree):
    beta = tree.xpath('//*[@id="quote-summary"]/div[2]/table/tbody/tr[2]/td[2]/span')[0].text
    #print 'beta = ', beta
    return beta

#Goes to the fidelity website, parses out the info we want and returns a dict
def StockParse():
    # Get Stock info
    page = requests.get('https://eresearch.fidelity.com/eresearch/gotoBL/fidelityTopOrders.jhtml', verify=False)
    tree = html.fromstring(page.content)

    #make arrays of length 30.  "1" is just initializing them all with the string "1".  This should not stay in the array.
    sTicker= ["1"] * 30
    sPriceChange = ["1"] * 30
    sBuy = ["1"] * 30
    sSell = ["1"] * 30
    sBuySellRatio = ["1"] * 30
    sCurPrice = ["1"] * 30
    sBeta = ["1"] * 30
    sPriceChangePercent = ["1"] *30
    sYahooPriceChange = ["1"] *30
    
    #Stock Lib is a dictionary.  It is how we return all of the values from this function.  Keep in mind it does not have labels.
    #Returns as #{{ticker1: price1, priceChange1, buy#1,sell#1,ratio#1},{ticker2: price2,priceChange2, buy#2,sell#2,ratio#2}, {3....}... }
    stockLib = {}

    #Start at i=1, end at 31.  (30 iterations)  Note: may give a syntax error in the morning b/c the website does not have 30 stocks listed
    for i in range (1,31):
        
        tickerListString = "//*[@id=\"topOrdersTable\"]/tbody/tr["
        tickerListString += str(i)
        tickerListString += "]/td[2]/span"

        priceChangeListString = "//*[@id=\"topOrdersTable\"]/tbody/tr["
        priceChangeListString += str(i)
        priceChangeListString += "]/td[4]/span"

        buyOrderListString = "//*[@id=\"topOrdersTable\"]/tbody/tr["
        buyOrderListString += str(i)
        buyOrderListString += "]/td[5]/span"
        
        sellOrderListString = "//*[@id=\"topOrdersTable\"]/tbody/tr["
        sellOrderListString += str(i)
        sellOrderListString += "]/td[7]/span"            

        #Actually get the fidelity data and put into arrays
        #i-1 because python arrays start at 0.
        sTicker[i-1] = tree.xpath(tickerListString)[0].text
        sPriceChange[i-1] = round(float((tree.xpath(priceChangeListString)[0].text).replace(",","")),2) #note, this is from fidelity and not updated always
        sBuy[i-1] = tree.xpath(buyOrderListString)[0].text
        sSell[i-1] = tree.xpath(sellOrderListString)[0].text
        
        #Yahoo Data
        treeYahoo = GetYahoo_MainPage(sTicker[i-1])
        
        sCurPrice[i-1] = float(GetYahooStock_Price(treeYahoo))
        sBeta[i-1] = GetYahooStock_Beta(treeYahoo)
        sYahooPriceChange[i-1] = float(GetYahooStock_PriceChange(treeYahoo)) #Note: formatting is wrong.
        
        #calculated data
        #Round to 2nd decimal space
        sBuySellRatio[i-1] = round((float(sBuy[i-1]) / (float(sBuy[i-1]) + float(sSell[i-1]))),2)
        sPriceChangePercent[i-1] = round(float(sYahooPriceChange[i-1] / sCurPrice[i-1])*100,2) #fidelity is not always updated.  Using Yahoo info instead
        
        #print 'ticker = ', sTicker[i-1]
        #print 'price = ', sPriceChangePercent[i-1]
        
        #Add it to the stock Dictionary
        stockLib[sTicker[i-1]] = [sCurPrice[i-1], sYahooPriceChange[i-1], sPriceChangePercent[i-1], sBuy[i-1], sSell[i-1], sBuySellRatio[i-1], sBeta[i-1]]
    
    #Debug Prints
    #print 'sTicker ', sTicker
    #print 'sPriceChange ', sPriceChange
    #print 'sBuy ', sBuy
    #print 'sSell ', sSell
    #print 'sBuySellRatio ', sBuySellRatio        
    #print 'sYahooPriceChange ', sYahooPriceChange
    
    #print 'stockLib ', stockLib
    
    return stockLib
    
    
    #For XPARSE strings uncomment this!
    #print 'tickerListString: ' tickerListString
    #print 'PriceListString: ' priceListString
    #print 'buyOrderListString: ' buyOrderListString
    #print 'sellOrderListString: ' sellOrderListString

        #Line for stock company name
#        Stocks = tree.xpath('//*[@id="topOrdersTable"]/tbody/tr[3]/td[3]')[0].text
#        print 'Stocks: ', Stocks
#        print("Done")



def main():

    #Stock Lib is a dictionary.  It is how we return all of the values from this function.  Keep in mind it does not have labels.
    #Returns as #{{ticker1: price1, buy#1,sell#1,ratio#1},{ticker2: price2, buy#2,sell#2,ratio#2}, {3....}... }
    #Will need to figure out how to add yahoo functions to it as well.  Also need to figure out how to make it look nice in excel
    stockLib = {}
    stockLib = StockParse()
    
    
    print 'Ticker: Price, PriceChange, price change %, buy amount, sell amount, ratio, beta'
    print 'stockLib in MAIN ', stockLib



if __name__ == '__main__':
    main()