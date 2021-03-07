# Get shares from finance.yahoo.com using VBA
This table gets shares from the finance.yahoo.com. The finance.yahoo.com API is used. No “SeleniumBasic” or other libraries are necessary. 
It is still at initial state (simple proof of concept) and will be tested / improved in the future. 
Note, that “investment horizon” in this table is just the time to calculate the performance. It will be renamed in the future. 
The table has following sheets:

## Shares
It contains the graphic and the possibilities to sort the shares. The updating of shares is made over the tickers in this table, so, if you added a new ticker, please add it first in “TimeSeries” and put here the reference to it as “=TimeSeries!$A$11”. Both “$” symbols are important for sorting, please do not forget them! The ticker is the only necessary parameter of the share.

## TimeSeries
Contains the open, high, low and close prices over one year. The average price is the average of open, high, low and close prices. This is the basis of the table. It still should be checked because, for example, for Bombardier the correct price is only at the end. 

## Currencies
The first step before updating the shares is to get the currencies because the currencies are necessary to calculate price in euro (sheet “Shares”, column “U”).

## Simulations
The simulations are done to adjust the investment horizon. The VBA sub “Model” contains the algorithm of investment. It Invests initial 10.000 (the currencies are not checked at this time) and invests them over the year at different “investment horizons” (1 to 30 days). This part hast to be adjusted because its performance is not as good as awaited.

## Proof
In row 21 the average performance over the year is calculated. One can see that the table, adjusted to investment horizon of 23 days gives approximately 500% per year, whereas the “standard” investment horizon of 3 months gives “only” 200%. The same shares that were the best at different time steps are marked green. It is quite unbelievable, but the best shares after 3 months coincide more than the best shares after 23 days (compare “E:F” and “O:P”)! 

## TODO:
In the future the following features will be implemented:
- The investment algorithm will be adjusted. (see sheet “Simulations”) I do not want to buy/sell the shares every day as predicted in sheet “Simulations”!
- The AI will be added. 
