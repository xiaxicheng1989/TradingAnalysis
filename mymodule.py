#import seaborn as sns
#import matplotlib.pyplot as plt
#import matplotlib.axes as ax
#import sklearn
#from sklearn.linear_model import LinearRegression
#from sklearn import datasets, linear_model
#from scipy.optimize import curve_fit
#import os
#import collections
#from statsmodels.stats.outliers_influence import summary_table


from pandas.tseries.offsets import *



#import sklearn
#from sklearn import datasets, linear_model
#from sklearn import datasets, linear_model
#import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from sklearn import datasets, linear_model
import xlwings as xw


def world():
    
    DataTrain = GetDataFrame("RawData_Train",1700,2000)
    DataTrain = filterData(DataTrain, "09/04/2019", "09/04/2019", "14:30", "21:00")
    DataValid = GetDataFrame("RawData_Train2",1700,2000)
    DataValid = filterData(DataValid, "09/05/2019", "09/05/2019", "14:30", "20:59")
    ###
    allData = pd.concat([DataTrain, DataValid]).dropna()
    ###
    
    #DataReturnTrain = getReturnTrain(DataTrain, 396+10, "MSFT US Equity-Open")
    
    #DataReturnValid = getReturnTrain(DataValid, 396+10, "MSFT US Equity-Open")
    
    lassoDF = getLassoValidDF(allData, 300+10, "MSFT US Equity-Open",0.01)
    
    peerDF = getCoefDF(allData,1, 300+10, "MSFT US Equity-Open",0.01)
    peerDF = getContr(peerDF).sort_values(by=["Contribution [%]"], ascending=False)
    
    outsht = xw.Book.caller().sheets["Results"]
    outsht.range('A1:F500').value = lassoDF
    peersht = xw.Book.caller().sheets["Peers"]
    peersht.range('A1:C100').value = peerDF
    #
    #peersht.range('A1:B100').number_format = '0.0'
    #test = xw.Book.caller().sheets["Sheet1"]
    #test.range("A2").value = 200
###############################################################################
def showLivePrediction(compName, alpha0):
    # prepare training + prediction data set
    trainData = prepDataSet()
    allData = createNewDF(trainData).dropna(axis=1,how="all")
    #get parameters
    mainsht = xw.Book.caller().sheets["Main"]
    compName0 = compName+'-Open'
    #alpha0 = mainsht.range('K2').value
    # creat lasso prediction dataframe
    lassoDF = getLassoValidDF(allData, len(trainData), compName0,alpha0)
    # print the dataframe
    outsht = xw.Book.caller().sheets["Results"]
    outsht.range('A1:F500').value = ""
    outsht.range('A1:F500').value = lassoDF
    
    
###############################################################################
def createNewDF(trainDF):
    # get parameters
    mainsht = xw.Book.caller().sheets["Main"]
    dateInterest = str(mainsht.range('I12').value[1:-1])
    dateInterest = str(mainsht.range('I12').value[1:-1])
    endTimeInterest = str(mainsht.range('I14').value[1:-1])
    
    trainData = prepDataSet().dropna(axis=1,how="all")
    liveData0 = GetDataFrame("DataValid",1700,2000)
    liveData = filterData(liveData0, dateInterest, dateInterest, endTimeInterest, "21:30")
    newDF = pd.concat([trainData, liveData],axis =0).fillna(value=0)
    
    return newDF
###############################################################################
def myPred():
    mainsht = xw.Book.caller().sheets["Main"]
    dateInterest = str(mainsht.range('I12').value[1:-1])
    endTimeInterest = str(mainsht.range('I14').value[1:-1])
    liveData0 = GetDataFrame("DataValid",1700,2000).fillna(value=0)
    liveData = filterData(liveData0, dateInterest, dateInterest, endTimeInterest, "21:30")
    
    coefListDF = mainsht.range("L1:M100").options(pd.DataFrame, index=True, numbers=float).value
    coefListDF = coefListDF[coefListDF.index.isna()==False]
    
    compName = mainsht.range('K1').value+'-Open'
    selCompList = coefListDF.index.tolist()
    datalist = np.array(liveData[compName]-liveData[compName][0])
    tes=np.array([np.sum([liveData[selCompList[i]][j]*coefListDF["Coef"][i] for i in range(len(selCompList)-1)]) for j in range(len(liveData))])
    
    resDF = pd.DataFrame()
    resDF["Data"] = datalist
    resDF["Basket"] = tes-tes[0]
    
    test = xw.Book.caller().sheets["Sheet1"]
    test.range('A1:B10').value = resDF
    
    
###############################################################################    
    
def GetPeerParameters(ticker, alpha0):
    # get the settings from the Main-tab.
    
    # initialise the Excel sheet
    mainsht = xw.Book.caller().sheets["Main"]
    # load the desired Name of the stock
    #compName0 = mainsht.range('K1').value+'-Open'
    compName = str(ticker)+'-Open'
    # load the predefined alpha value
    #alpha = mainsht.range('K2').value
    # clear the output space
    mainsht.range('L1:N200').value = ""
    
    # loading the master datafrane and erase columns which dont have any values.
    allData = prepDataSet().fillna(method="ffill").fillna(method="bfill").dropna(axis=1,how="all")
    
    # collect Lasso training output
    peerDFsummary = getCoefDF(allData,1, len(allData), compName, alpha0)
    # calculating the contribution of the peers
    peerDF = getContr(peerDFsummary[0]).sort_values(by=["Contribution [%]"], ascending=False)
    
    # Output coefficient dataframe into the space cleaned before
    mainsht.range('L1:N100').value = peerDF
    
    # Output the Lasso Model calculated training data 
    # initialise the tab for training results       
    trainressht = xw.Book.caller().sheets["Result_Train"]
    # clear the space
    trainressht.range('A1:E3000').value = ""
    # Saving
    trainressht.range('A1:E10').value = peerDFsummary[1]
    return peerDF.index
    
  # =IFERROR(VLOOKUP(K1,Sheet1!A2:B6,2,FALSE),0.01)  
###############################################################################
def prepDataSet():
    # get the settings from the Main-tab.
    mainsht = xw.Book.caller().sheets["Main"]
    startDate = str(mainsht.range('I1').value[1:-1])
    endDate = str(mainsht.range('I2').value[1:-1])
    startTime = str(mainsht.range('I4').value[1:-1])
    endTime = str(mainsht.range('I5').value[1:-1])
    dateInterest = str(mainsht.range('I12').value[1:-1])
    stTimeInterest = str(mainsht.range('I13').value[1:-1])
    endTimeInterest = str(mainsht.range('I14').value[1:-1])
   
    # load RawData_Train tab into DataFrame
    DataTrain = GetDataFrame("RawData_Train",5000,2000)
    # xlwings doesnt know when to stop loading. Therefore, i have loaded more
    # than needed into the dataframe. In order not to run into trouble with data
    # filling later, I delete everything with doesnt have an index.
    DataTrain = DataTrain[DataTrain.index.isna() == False]
    # the data can be constructed of multiple days. Here i only take the time 
    # when the market is open. Then I apply data filling on empty cells.
    DataTrain = filterData(DataTrain, startDate, dateInterest, startTime, endTime)
    
    # Load RawData_Train2 into dataframe
    DataTrain2 = GetDataFrame("RawData_Train2",5000,2000)
    DataTrain2 = DataTrain2[DataTrain2.index.isna() == False]
    DataTrain2 = filterData(DataTrain2, dateInterest, dateInterest, stTimeInterest,endTimeInterest)
    
    # combine both dataframes into the same one
    allData = pd.concat([DataTrain, DataTrain2])
    
    return allData

###############################################################################    
# This makes the DataFrame
def GetDataFrame(sheetname,N,M):
    # adressing the desired Excel sheet
    target = xw.Book.caller().sheets[sheetname] 
    # import the defined range into dataframe
    Data = target.range((1,1),(N,M)).options(pd.DataFrame, index=False, numbers=float).value
    # Delete empty rows and columns
    Data = Data.dropna(how='all',axis=1)
    Data = Data[Data.index.isna() == False]

    
    # Condense the loaded dataframe, as the original data has repeated time columns
    # and empty coloums
    #
    # Aquire all the Stock names
    labelList = [str(i)+"-Open" for i in Data.iloc[1,:].tolist()[0::2]]
    # Aquire only the stock prices
    DataPart = Data.iloc[3:,1::2]
    # Label the prices
    DataPart.columns = labelList
    # Define a common time index
    DataPart["Dates"] = pd.to_datetime(Data.iloc[3:,0])
    DataPart = DataPart.set_index("Dates")
    # Apply forward-filling and then backward-filling
    Data = DataPart.fillna(method="ffill").fillna(method="bfill")
  
    return Data


def filterData(dataDF, startDate, endDate, startTime, endTime):
    # If-condition distinguishes the case of single day history data or
    # multiple days of history data
    if startDate == endDate:
        stHour=startDate+" "+startTime
        clHour=startDate+" "+endTime
        # For single day data, I only select data when the market is open
        dataDF=dataDF[(dataDF.index>=stHour)&(dataDF.index<clHour)]
    else:
        # For multiple days:
        # I create lists of market opening times and closing times
        stHour=pd.date_range(startDate+" "+startTime, endDate)
        clHour=pd.date_range(startDate+" "+endTime, endDate)
        # Creat single day dataframe within the opening hours
        resDF = []
        for i in range(len(stHour)):
            tempDF=dataDF[(dataDF.index>=stHour[i])&(dataDF.index<clHour[i])]
            resDF.append(tempDF)
        # combine the multi day data
        dataDF = pd.concat(resDF)
    return dataDF

###############################################################################
# This is all lasso stuff
###############################################################################

def getLassoValidDF(data, nrTrain, stockname, alph):
    # convert into returns
    validdata=getReturnValid(data, nrTrain, stockname)
    # predition variables
    Xvalid = validdata[0].fillna(value=0)
    
    # creat a output dataframe
    Yvalid=validdata[1].dropna(axis='columns')
    fitfct = getLassoFit(data, nrTrain, stockname, alph)
    Yvalid["Basket"] = fitfct.predict(Xvalid)-fitfct.predict(Xvalid)[0]
    Yvalid["vsBasket"] = Yvalid["Data"]-Yvalid["Basket"]
    Yvalid["MSE"] = Yvalid["vsBasket"]
    return Yvalid

def getReturnValid(data, nrTrain, stockname):
    raw = prepValidData(data, nrTrain, stockname)
    for j in range(2):
        rawcol= raw[j].columns
        for i in rawcol:
            val = raw[j][i].values[0]
            raw[j][i] =(raw[j][i]-val)*10000.0/val
        raw[j].fillna(method="bfill").fillna(method="ffill").dropna(axis='columns')
    return raw

def prepValidData(data, nrTrain, stockname):
    valid = data[nrTrain:]
    Xvalid = valid.drop(stockname, axis=1)
    Yvalid = pd.DataFrame()
    Yvalid["Data"]=valid[stockname]
    return [Xvalid, Yvalid]

def getLassoFit(data, nrTrain, stockname, alph):
    # prepare the raw data into return data
    train = getReturnTrain(data, nrTrain, stockname)
    # lasso
    lml = linear_model.Lasso(alpha=alph, normalize=True,tol=0.00001, max_iter=500)
    est2=lml.fit(train[0].dropna(axis='columns'),train[1].dropna(axis='columns'))
    # outputs an Lasso object
    return est2

def getReturnTrain(data, nrTrain, stockname):
    # separate data into training variables and target set
    raw = prepTrainData(data, nrTrain, stockname)
    # The outcome is a list of 2. Here we convert both into returns.
    for j in range(2):
        rawcol= raw[j].columns
        for i in rawcol:
            val = raw[j][i].values[0]
            raw[j][i] = (raw[j][i]-val)*10000.0/val
        # filling 
        raw[j].fillna(method="bfill").fillna(method="ffill").dropna(axis='columns')    
    return raw

def prepTrainData(data, nrTrain, stockname):
    train = data[:nrTrain]
    Xtrain = train.drop(stockname, axis=1)
    Ytrain = pd.DataFrame()
    Ytrain["Data"]=train[stockname]
    return [Xtrain, Ytrain]

###############################################################################
# Peer names
###############################################################################

def getCoefDF(data, avg, nrTrain, stockname, alph):
    # Converting the open price into returns and separate the company of interest 
    # from the rest.
    traindata=getReturnTrain(data, nrTrain, stockname)
    
    # define the training data
    Xtrain = traindata[0]
    # get lasso coefficients
    fitfct = getLassoFit(data, int(nrTrain/avg), stockname, alph)
    # construct the dataframe for the lasso coefficients with the peer names
    # as index
    coefDF = pd.DataFrame(index = Xtrain.columns)
    # assign the index with coefficients
    coefDF["Coef"] = fitfct.coef_
    
    # Construct dataframe for the traingin results
    # History data of the company of interest
    trainResDF = traindata[1]
    # Model outcom on describing the past data
    trainResDF["Basket"] = fitfct.predict(Xtrain)
    # Form vsBasket values
    trainResDF["vsBasket"] = trainResDF["Data"]-trainResDF["Basket"]
    trainResDF["MSE"] = (trainResDF["vsBasket"])**2
    
    # output coefficient DF and training results
    return [coefDF, trainResDF]
   
def getContr(peerDFdata):
    # create a dataframe only for peers with non-zero coefficients
    filtered =peerDFdata[peerDFdata["Coef"] != 0]
    # Calculating the contribution
    norm = np.sum(filtered.values**2)
    contributionDF = filtered
    contributionDF["Contribution [%]"] = 100*filtered.values**2/norm
    return contributionDF
