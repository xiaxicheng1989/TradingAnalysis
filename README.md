# Trading Analysis - Basket analysis
This is the work, I have done for real time trading analysis.

I combine VBA and python for the project, where VBA is automated to extract data from Bloomberg terminal and Python is used for analysis.

This is the outcome on the real-time trading analysis on the European and American stocks. Here, I use linear dimensionality reduction (lasso) to estimate the price return of any equity based on the peer share price. This can be used to analyse subjective market participation.

The structure of the programme is described by:
<img src="https://github.com/xiaxicheng1989/TradingAnalysis/blob/master/plots/schematic.png" width="50%">
The data will be pulled from excel using VBA (DBH-functions) and passed on to Python for simulation and calculation via [xlwings](https://www.xlwings.org/) - excel macro . The results will be passed back to Excel for representation.

