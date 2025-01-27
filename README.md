# Trading Analysis - Basket analysis
This is the work, I have done for real time trading analysis.

I combine VBA and python for the project, where VBA is automated to extract data from Bloomberg terminal and Python is used for analysis.

This is the outcome on the real-time trading analysis on the European and American stocks. Here, I use linear dimensionality reduction (lasso) to estimate the price return of any equity based on the peer share price. This can be used to analyse subjective market participation.

The structure of the programme (VBATesting.xlsx) is described by:  

<img src="https://github.com/xiaxicheng1989/TradingAnalysis/blob/master/plots/schematic.png" width="50%">

The data will be pulled from excel using VBA (DBH-functions) and passed on to Python for simulation and calculation via [xlwings](https://www.xlwings.org/) - excel macro . The results will be passed back to Excel for representation.

The structure of the programme is illustrated here:

<img src="https://github.com/xiaxicheng1989/TradingAnalysis/blob/master/plots/programmeSchematic.png" width="50%">

Green is Excel/VBA. Blue is python. <…> are all vba codes. Some of them triggers the python script. The codes are all global functions. Object orientated programming could be more efficient though.

## VBA:
VBA triggers python by running “RunPython”-function of xlwings library (see VBA environment.) It is an one-liner. As the input is a string, it is quite hard to call a function in python, which takes a argument. This will be useful if we get it working. According to xlwings’ helpdesk, “RunPython” can be made such way that it takes an argument but not convenient. One needs to combine strings including the “ ‘ ”.  
Also, RunPython doesn’t return values to pass on. To do this it is suggested to use User Defined Functions (UDFs).

### Issue of data consumption on Bloomberg:
Bloomberg restricts a user’s daily data usage to be 500 000 hits and monthly cap will be calculated using an algorithmic weighting system, which depends on the fields. There are no further information on how they calculate that. According to the agent, pulling past history data only counts as 1 hit, if it is applied to one field.

#### My understanding: 
This seems to make sense. As Bloomberg doesn’t own any raw data, they are not allowed to charge us for that, but only for the service. Hence, they charge for the number of times of entering the function “DBH” or so on, which is essentially their service. This is also why they introduced the daily/monthly cap to limit the maximal number of times for ppl to access their server. Any processed data by Bloomberg seems to require the usage of their server, therefore the weird weighting system. This means we are good with the data consumption. A rough estimate for our daily data usage is given below with the following assumptions:
- three markets
- 500 universe peers in each market (more than we actually need)
- One security type per peer
- DBH function for each peer will be executed twice (once for the main training data and once for the extra minutes) 
- 20 equities to be traded per market per day 
- 3 Lasso parameters for each trade equity
- 50 selected stocks for each lasso setting
- Refresh the live data of the each selected stock 30 times per day.

This makes 3 x (500 x 1 x 2 + 20 x 3 x 50 x 30) = 273 000 hits. This turns out to be quite large, but we are under daily cap.

## Python:
VBA mainly calls two python functions:  <code>GetPeerParameters</code> and <code>showLivePrediction</code>. The structure of the python part of the software is illustrated below:

<img src="https://github.com/xiaxicheng1989/TradingAnalysis/blob/master/plots/pythonstructure.png" width="60%">

More detail on the function description can be found in "FunctionDescription.docx"

# Backbone for the programme
The development of this VBA is purely based on the master notebook, which was used to explore the equity data. (Northamerica.ipynb)

Briefly summarised:  
- Data cleaning: Equity data downlaoded from bloomberg arent nessesarily clean, hence, data needs to be filled, which in turn affects prediciton outcome. After going through all data of 500 companies, we can illustrate the number of companies with maximal percentage of missing data as below:

<img src="https://github.com/xiaxicheng1989/TradingAnalysis/blob/master/plots/missingdata.png" width="40%">

Before further moving on to modeling, the share prices are normalised.

- Modeling: Modeling is mainly done using Lasso. Here I optimise:
  1. Lasso alpha
  2. Length of the traning data 
  3. Length of the testing data
  4. Time bin of average to reduce the noise
 An example is shown below:
 
<img src="https://github.com/xiaxicheng1989/TradingAnalysis/blob/master/plots/example.png" width="80%">

(The top graph show the training set in orange and the target in blue. The second graph shows the prediction of the next day's price in orange, compared with the actual return in basis points in blue. The final graph shows the difference in the prediction and the real price. The color coding illustrates the level of mean reversion.)


To validate the model, I not only use the MSE for the validation set. As mean reversion is demanded, I look at the difference between the predicted price and the real price. A minimum in the sum of the difference and the peak-peak(min/max) would give us a good indication of correct prediction. These has been applied on two liquid and two iliquid equities for testing in the notebook.  
 
 
# Outlook
To be able to correctly judge if market participation becomes significant, one could construct indicator based on the min-max of the mean reversion graph. As expected, the trigger level will be different depending on the liquidity of the equity, which need to be evaluated using past data for each company.

Inaddition, the codes were writting using 80:20 rule. The focus was on understanding the driving force. Hence, the progromme is fairly slow and inefficient.
