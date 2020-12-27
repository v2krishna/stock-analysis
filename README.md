# Stock Analysis

Perform stock analysis for the provided companies based on the year selected. Also refactor the code wherever possible to run the stock analysis in very less time.

## Overview of Project
Perform Stock analysis for 12 companies containing the data for 2017 and 2018. Scope of the project is to get the percentage change of the prices of each company. Percentage change is calculated based on the price at the start of the year and price at the end of the year. By doing the yearly returns of all the companies, will better help me to identify in which stock need to invest in the future.

### Purpose
Stock Performance of each ticker symbol for selected year. At the same time refactor the initial VB Script which was taking more time somewhere between 40 to 50 seconds for one year, so it takes less time to execute without compromising on the data accuracy.


## Stock Performance and Measuring the Refacored VBA
For calculating the stock performance of the prices or returns, we need to first get all the stocks symbols , sort the data by symbol and date in ascending order and loop the data for each ticker symbol, calculate the starting price and ending prices of the ticker within a year.

### Stock Performance
![2017_StockPerformance](/resources/2017_StockPerformance.png) ![2018_StockPerformance](/resources/2018_StockPerformance.png)<br/>
The above stock performances are for 2017 and 2018 respectively. In 2017 though DQ has performed extraordinarily with 199% returns, but in 2018, it was down by 62%, based on this returns, would like to rethink on investing in DQ. By comparing the returns of 2017 and 2018 , ENPH and RUN are the two stocks which is giving is consistent returns, will look into those stocks for investing.

### Refactor Measurements
![2017_BeforeRefactoring](/resources/2017_BeforeRefactoring.png)  ![VBA_Challenge_2017](/resources/VBA_Challenge_2017.png) </br>
Based on the above results, the image on the left is taken when stock analysis is performed on the inital VB code for 2017 year, the time it has taken more than 36 seconds whereas the image on the right is taken after refactoring the code for 2017, duration for executing is less than 0.2 seconds. Similarly the same has been observed for 2018, before refactoring the code, its taking around 41 seconds to complete the process, where as after refactoring the code, its less than 0.4secs. <br/>
![2018_BeforeRefactoring](/resources/2018_BeforeRefactoring.png)  ![VBA_Challenge_2018](/resources/VBA_Challenge_2018.png) <br/>. 
Also made sure that the data is same before and after the refactoring the code.
![2017_StockPerformance](/resources/2018_StockPerformance.png) ![2018_StockPerformance](/resources/2018_StockPerformance.png) <br/>
