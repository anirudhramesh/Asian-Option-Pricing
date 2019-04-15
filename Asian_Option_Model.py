# Asian option pricing model implemented using monte carlo simulation (normal and antithetic)
# and compared against vanilla Black Scholes model

import math
import numpy as np
from scipy.stats import norm
import random
import xlrd as excel
from xlwt import Workbook
import matplotlib.pyplot as plotter
import matplotlib.dates as mdates
from datetime import datetime

def price_asian_option_vanilla(S0,rD, rF, diffrates,sigma,T,n,m,K):
    callPrice=0
    putPrice=0
    mu, sd = 0, 1
    stdev=[]
    for i in range(n):
        sumvalue=0
        product=1
        for j in range(m):
            X = np.random.normal(mu, sd)
            ST=S0*math.exp((diffrates-(sigma**2/2))*T+sigma*(((j*T)/m)**0.5)*X)
            # sumvalue+=ST/(m)
            product*=ST**(1/(m))
        callPrice+=max(0,product*math.exp(-rF*T)-K*math.exp(-rD*T))/n
        stdev.append(product)
        # putPrice+=max(0,K-product)
        # putPrice += max(0, -sumvalue * math.exp(-rF * T) + K * math.exp(-rD * T)) / n
        #total+=max(0,sumvalue/(m-1) - K)/n
    putPrice=max(0,callPrice+K*math.exp(-rD*T)-S0)
    stdev=np.std(stdev)/math.sqrt(n)
    return putPrice, stdev


def price_asian_option_antithetic(S0, rD, rF, diffrates, sigma, T, n, m, K):
    callPrice = 0
    putPrice = 0
    mu, sd = 0, 1
    product = 1
    stdev=[]
    for i in range(n):
        sumvalue = 0
        product = 1
        for j in range(m):
            X = np.random.normal(mu, sd)
            ST = S0 * math.exp((diffrates - (sigma ** 2 / 2)) * T + sigma * (((j * T) / m) ** 0.5) * X)
            ST1 = S0 * math.exp((diffrates - (sigma ** 2 / 2)) * T + sigma * (((j * T) / m) ** 0.5) * -X)  # antithetic variate
            # sumvalue += (lists + lists1) / 2
            product *= ((ST + ST1) / 2)**(1/m)
        callPrice += max(0,product*math.exp(-rF*T) -K*math.exp(-rD*T)) / n
        stdev.append(product)
        # putPrice += max(0, -product * math.exp(-rF * T) + K * math.exp(-rD * T)) / n
        # total+=max(0,sumvalue/(m-1) - K)/n
    putPrice = max(0,callPrice + K * math.exp(-rD * T) - S0)
    stdev=np.std(stdev)/math.sqrt(n)
    return putPrice, stdev

def price_option_BS(S0, rD, rF, diffrates, sigma, T, n, K):
    d1=(math.log(S0/K)+(diffrates+(sigma**2)/2)*T)/(sigma*math.sqrt(T))
    d2 = (math.log(S0 / K) + (diffrates - (sigma ** 2) / 2) * T) / (sigma * math.sqrt(T))
    callPrice=S0*norm.cdf(d1)-K*math.exp(-rF*T)*norm.cdf(d2)
    putPrice = max(0, callPrice + K * math.exp(-rD * T) - S0)
    # print(str(S0) + str(K) + str(r) + str(sigma))
    return putPrice

# The use case is pricing a currency option for the USDINR pair
def main():
    usdinr=[]
    usdinr_change=[]
    dates=[]
    diffrates=[]
    ratesD=[]
    ratesF=[]
    
    # expiry in 12 months
    T=12

    wb=excel.open_workbook(filename='India_data.xlsx')
    wb1=Workbook()
    ws1=wb1.add_sheet('Prices')
    
    # code outputs tons of data to an excel spreadsheet including prices using
    # 1. Vanilla monte carlo
    # 2. Antithetic monte carlo
    # 3. Black Scholes Model
    # and also outputs the standard error of these price estimations
    
    # We simulate 4 different paths with strike prices a few std's away from the actual prevailing price on the date.
    # The labels below aren't accurate. We settled on different sigma values in the final simulation.
    
    ws1.write(0, 0, 'Date')
    ws1.write(0, 1, 'USDINR')
    ws1.write(0, 2, 'Vanilla K=-1sigma')
    ws1.write(0, 3, 'Vanilla K=-0.5sigma')
    ws1.write(0, 4, 'Vanilla K=+0.5sigma')
    ws1.write(0, 5, 'Vanilla K=+1sigma')
    ws1.write(0, 6, 'Antithetic K=-1sigma')
    ws1.write(0, 7, 'Antithetic K=-0.5sigma')
    ws1.write(0, 8, 'Antithetic K=+0.5sigma')
    ws1.write(0, 9, 'Antithetic K=+1sigma')
    ws1.write(0, 10, 'BS K=-1sigma')
    ws1.write(0, 11, 'BS K=-0.5sigma')
    ws1.write(0, 12, 'BS K=+0.5sigma')
    ws1.write(0, 13, 'BS K=+1sigma')
    ws1.write(0, 14, 'Strike Price K=-1sigma')
    ws1.write(0, 15, 'Strike Price K=-0.5sigma')
    ws1.write(0, 16, 'Strike Price K=+0.5sigma')
    ws1.write(0, 17, 'Strike Price K=+1sigma')
    ws1.write(0, 18, 'Std Error Vanilla K=-1sigma')
    ws1.write(0, 19, 'Std Error Vanilla K=-0.5sigma')
    ws1.write(0, 20, 'Std Error Vanilla K=+0.5sigma')
    ws1.write(0, 21, 'Std Error Vanilla K=+1sigma')
    ws1.write(0, 22, 'Std Error Antithetic K=-1sigma')
    ws1.write(0, 23, 'Std Error Antithetic K=-0.5sigma')
    ws1.write(0, 24, 'Std Error Antithetic K=+0.5sigma')
    ws1.write(0, 25, 'Std Error Antithetic K=+1sigma')

    # cell co-ordinates work like array indexes within the spreadsheet
    for i in range(2,121):
        single_usd_inr=wb.sheet_by_name("Sheet1").cell_value(i,3)
        usdinr.append(single_usd_inr)
        usdinr_change.append(math.log(1+wb.sheet_by_name("Sheet1").cell_value(i, 4)))
        single_date=wb.sheet_by_name("Sheet1").cell_value(i, 0)
        ws1.write(i, 0, single_date)
        ws1.write(i, 1, single_usd_inr)
        dates.append(datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(single_date) - 2))
        ratesD.append(math.log(1 + wb.sheet_by_name("Sheet1").cell_value(i, 1) / 100) / 12)
        ratesF.append(math.log(1 + wb.sheet_by_name("Sheet1").cell_value(i, 9) / 100) / 12)
        diffrates.append(ratesD[i-2]-ratesF[i-2])

    optionPricesVanilla=[]
    optionPricesAntithetic=[]
    optionPricesBS=[]
    for i in range(12, len(dates)-T):
        sigma1 = np.std(usdinr_change[i-12:i-1])
        sigma = np.std(usdinr[i - 12:i - 1])
        mean = np.average(usdinr_change[i-12:i-1])
        
        # Strike prices are decided as +/- [0.75, 0.25] sigma from spot price.
        # Since this is a simulation, we calculate the option price through simulated paths and compare it against the
        # actual strike price during the selected expiry date to decide if the option is exercised or not.
        K = [usdinr[i] - 0.75 * sigma * math.sqrt(T), usdinr[i] - 0.25 * sigma * math.sqrt(T), usdinr[i] + 0.25 * sigma * math.sqrt(T), usdinr[i] + 0.75 * sigma * math.sqrt(T)]
        ws1.write(i, 14, K[0])
        ws1.write(i, 15, K[1])
        ws1.write(i, 16, K[2])
        ws1.write(i, 17, K[3])

        optionPrice = []
        stderr = []
        for k in K:
            optionPrice_single, stderr_single = price_asian_option_vanilla(usdinr[i], ratesD[i], ratesF[i], diffrates[i], sigma1, T, 100, 120, k)
            optionPrice.append(optionPrice_single)
            stderr.append(stderr_single)
        optionPricesVanilla.append(optionPrice)
        ws1.write(i, 2, optionPrice[0])
        ws1.write(i, 3, optionPrice[1])
        ws1.write(i, 4, optionPrice[2])
        ws1.write(i, 5, optionPrice[3])
        ws1.write(i, 18, stderr[0])
        ws1.write(i, 19, stderr[1])
        ws1.write(i, 20, stderr[2])
        ws1.write(i, 21, stderr[3])

        optionPrice = []
        stderr = []
        for k in K:
            optionPrice_single, stderr_single = price_asian_option_antithetic(usdinr[i], ratesD[i], ratesF[i], diffrates[i], sigma1, T, 100, 120, k)
            optionPrice.append(optionPrice_single)
            stderr.append(stderr_single)
        optionPricesAntithetic.append(optionPrice)
        ws1.write(i, 6, optionPrice[0])
        ws1.write(i, 7, optionPrice[1])
        ws1.write(i, 8, optionPrice[2])
        ws1.write(i, 9, optionPrice[3])
        ws1.write(i, 22, stderr[0])
        ws1.write(i, 23, stderr[1])
        ws1.write(i, 24, stderr[2])
        ws1.write(i, 25, stderr[3])

        optionPrice=[price_option_BS(usdinr[i], ratesD[i], ratesF[i], diffrates[i], sigma1, T, 100, k) for k in K]
        optionPricesBS.append(optionPrice)
        ws1.write(i, 10, optionPrice[0])
        ws1.write(i, 11, optionPrice[1])
        ws1.write(i, 12, optionPrice[2])
        ws1.write(i, 13, optionPrice[3])

    optionPricesVanilla=np.array(optionPricesVanilla)
    optionPricesAntithetic = np.array(optionPricesAntithetic)
    optionPricesBS = np.array(optionPricesBS)

    wb1.save('Option Prices.xls')

    # finally, we plot all the different option prices from our simulations.
    axis1= plotter.subplot(2,2,1)
    axis1.plot_date(dates[12:len(dates)-T], optionPricesVanilla[:,0], label="-2sigma", linestyle='solid', marker=',')
    axis1.plot_date(dates[12:len(dates)-T], optionPricesVanilla[:,1], label="-1sigma", linestyle='solid', marker=',')
    axis1.plot_date(dates[12:len(dates)-T], optionPricesVanilla[:, 2], label="+1sigma", linestyle='solid', marker=',')
    axis1.plot_date(dates[12:len(dates)-T], optionPricesVanilla[:, 3], label="+2sigma", linestyle='solid', marker=',')
    axis1.set_ylabel('Option Prices')
    axis1.legend()
    axis2=axis1.twinx()
    axis2.plot_date(dates, usdinr, label='USDINR', linestyle='solid', marker=',')
    axis2.set_ylabel('USD / INR Prices')
    axis2.legend()
    plotter.title('Vanilla Monte Carlo Simulation')

    axis1 = plotter.subplot(2, 2, 2)
    axis1.plot_date(dates[12:len(dates) - T], optionPricesBS[:, 0], label="-2sigma", linestyle='solid', marker=',')
    axis1.plot_date(dates[12:len(dates) - T], optionPricesBS[:, 1], label="-1sigma", linestyle='solid', marker=',')
    axis1.plot_date(dates[12:len(dates) - T], optionPricesBS[:, 2], label="+1sigma", linestyle='solid', marker=',')
    axis1.plot_date(dates[12:len(dates) - T], optionPricesBS[:, 3], label="+2sigma", linestyle='solid', marker=',')
    axis1.set_ylabel('Option Prices')
    axis1.legend()
    # axis2 = axis1.twinx()
    # axis2.plot_date(dates, usdinr, label='USDINR', linestyle='solid', marker=',')  # , secondary_y=True)
    # axis2.set_ylabel('USD / INR Prices')
    # axis2.legend()
    plotter.title('Black Scholes Calculation')

    plotter.subplot(2, 2, 3)
    plotter.plot_date(dates[12:len(dates) - T], optionPricesAntithetic[:, 0], label="-2sigma", linestyle='solid', marker=',')
    plotter.plot_date(dates[12:len(dates) - T], optionPricesAntithetic[:, 1], label="-1sigma", linestyle='solid', marker=',')
    plotter.plot_date(dates[12:len(dates) - T], optionPricesAntithetic[:, 2], label="+1sigma", linestyle='solid', marker=',')
    plotter.plot_date(dates[12:len(dates) - T], optionPricesAntithetic[:, 3], label="+2sigma", linestyle='solid', marker=',')
    plotter.legend()
    plotter.title('Antithetic Monte Carlo Simulation')
    plotter.ylabel('Option Prices')
    
    plotter.show()


if(__name__=="__main__"):
    main()
