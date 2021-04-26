import os
import openpyxl as op
import pandas as pd
from datetime import datetime
import pytz
import time
import sys


def get_url(stock,start_date,end_date):
    url="https://query1.finance.yahoo.com/v7/finance/download/{}?period1={}&period2={}&interval=1d&events=history&includeAdjustedClose=true".format(stock,start_date,end_date)
    return url


def get_stock_price(stocklist):
    #This to create new file every time

    start_date = "1514725200" #Unix for 01/01/2018 00:00:00 time stamp    
    #current_time = datetime.now(pytz.timezone("GMT")) #Change current GMT+11 Sydney time to GMT+0, no need anymore
    current_time = datetime.now()
    end_date = int(time.mktime(current_time.timetuple())) #change to Unix time stamp

    filename = "AllShares.xlsx"
    with pd.ExcelWriter(filename, engine = "openpyxl" ,options ={"in_memory": True}) as writer:
        for stock in stocklist:
            #Get stock data           
            symbol = stock+".AX"  #to specify ASX stock
            url = get_url(symbol, start_date, end_date)

            #clean data
            stock_data = pd.read_csv(url) #get stock data based on url
            stock_data.sort_index(ascending=False, inplace = True)

            stock_data["Date"] = pd.to_datetime(stock_data["Date"], format ="%Y-%m-%d")#Change string type to datetime type
            stock_data["Date"]= stock_data["Date"].dt.strftime("%d-%m-%Y")#change to dd-mm-yyyy format

            data_col =["Open","High","Low","Close","Adj Close"]
            stock_data[data_col] = stock_data[data_col].round(3)
            stock_data["ShareID"] = stock
                       
            stock_data.to_excel(writer, sheet_name =stock ,index = False, header = True)
                


if __name__ == '__main__':
    stocklist = ["CBA","WBC","ANZ","NAB"]
    get_stock_price(stocklist)
    print ("Completed")

