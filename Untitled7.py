
# coding: utf-8

# In[ ]:


import numpy as np
import math
import os.path
from xlwt import Workbook
import time

def price_option_put_BS(S0, r, sigma, T, n, K):
    sumvalue = 0
    for i in range(n):
        X=np.random.normal(0, 1)
        #print(X)
        ST=S0*math.exp((r-(sigma**2)/2)*T+sigma*math.sqrt(T)*X)
        #print(ST)
        sumvalue += max(K-ST,0)
    putPrice=math.exp(-r*T)*sumvalue/(n-1)
    #d1=(math.log(S0/K)+(r+(sigma**2)/2)*T)/(sigma*math.sqrt(T))
    #d2 = (math.log(S0 / K) + (r - (sigma ** 2) / 2) * T) / (sigma * math.sqrt(T))
   # putPrice=S0*norm.cdf(d1)-K*math.exp(-r*T)*norm.cdf(d2)
   # print(str(S0) + str(K) + str(r) + str(sigma))
    return putPrice


def price_asian_option(S0, r, sigma, T, n, m, K=90):
    total = 0
    for i in range(n):
        sumvalue = 0
        for j in range(m):
            mu, sd = 0, 1
            X = np.random.normal(mu, sd)
            lists = S0 * math.exp((r - (sigma ** 2 / 2)) * T + sigma * (((j * T) / m) ** 0.5) * X)
            # print(lists)
            sumvalue += lists
        # print(sumvalue)
        total += max(0, (sumvalue / (m - 1)) - K)
        # print(total)

    return math.exp(-r * T) * (total / n)


price_option_put_BS(100, 0.06, 0.048188, 1, 100, 70)

import math
import numpy as np
import random
import xlrd as excel
import matplotlib.pyplot as plotter
import matplotlib.dates as mdates
from datetime import datetime
from xlwt import Workbook


def price_asian_option(S0, r, sigma, T, n, m, K=90):
    total = 0
    mu, sd = 0, 1
    product = 1
    for i in range(n):
        sumvalue = 0
        product = 1
        for j in range(m):
            X = np.random.normal(mu, sd)
            lists = S0 * math.exp((r - (sigma ** 2 / 2)) * T + sigma * (((j * T) / m) ** 0.5) * X)
            sumvalue += lists
            product *= lists
        total += max(0, K - product ** (1 / m)) / n
        # total+=max(0,sumvalue/(m-1) - K)/n

    return math.exp(-r * T) * total


def floatHourToTime(fh):
    h, r = divmod(fh, 1)
    m, r = divmod(r * 60, 1)
    return (
        int(h),
        int(m),
        int(r * 60),
    )


usdinr = []
r=[]
ra=[]
dates = []
rates = []
wb = excel.open_workbook("/Users/nataliejin/Downloads/housing price index (2).xls")
wb1=Workbook()
ws1=wb1.add_sheet('Option Prices')

ws1.write(0,0,'Date')
ws1.write(0,1,'K=-2sigma')
ws1.write(0,2,'K=-1sigma')
ws1.write(0,3,'K=+1sigma')
ws1.write(0,4,'K=+2sigma')

# cell co-ordinates work like array indexes
for i in range(0, 93):
    usdinr.append(wb.sheet_by_name("Sheet1").cell_value(i, 6))
    r.append(wb.sheet_by_name("Sheet1").cell_value(i, 4))
    ra.append(wb.sheet_by_name("Sheet1").cell_value(i, 1))
    # dates.append(excel.xldate_as_tuple(wb.sheet_by_name("Sheet1").cell_value(i, 0),wb.datemode))
    # dates.append(wb.sheet_by_name("Sheet1").cell_value(i, 0))
    cell = wb.sheet_by_name("Sheet1").cell_value(i, 0)
    dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(cell) - 2)
    hour, minute, second = floatHourToTime(cell % 1)
    dt = dt.replace(hour=hour, minute=minute, second=second)
    dates.append(dt)
    # dates=datetime.datetime(*excel.xldate_as_tuple(dates, wb.datemode))
    rates.append(wb.sheet_by_name("Sheet1").cell_value(i, 1))
    # print(dates)
optionPrices = []
for i in range(12, len(dates) - 4):
    sigma = np.std(ra[i-12:i-1])
    # print(sigma)
    K = [90 - 2 * sigma, 90 - sigma, 90 + sigma, 90 + 2 * sigma]
    optionPrices.append([price_asian_option(100, r[i] , sigma, 1, 4, 1000, k) for k in K])
    ws1.write(i-11, 1, optionPrices[i-12][0])
    ws1.write(i - 11, 2, optionPrices[i - 12][1])
    ws1.write(i - 11, 3, optionPrices[i - 12][2])
    ws1.write(i - 11, 4, optionPrices[i - 12][3])

optionPrices = np.array(optionPrices)
# print(optionPrices)
fig = plotter.figure()
plotter.plot(dates[12:len(dates) - 4], optionPrices[:, 0], label="-2sigma")
plotter.plot(dates[12:len(dates) - 4], optionPrices[:, 1], label="-1sigma")
plotter.plot(dates[12:len(dates) - 4], optionPrices[:, 2], label="+1sigma")
plotter.plot(dates[12:len(dates) - 4], optionPrices[:, 3], label="+2sigma")
plotter.legend()
# plotter.xaxis.set_major_formatter(mdates.datetimefmt("%d-%m-%y"))
plotter.show()


fig.savefig('blacksholes.png')
wb1.save('Option Prices.xls'

