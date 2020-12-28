# Stock Analysis

Perform stock analysis for the provided companies based on the year selected at the sametime refactor the code wherever possible to run the stock analysis in very less time.

## Overview of Project
Perform Stock analysis for 12 companies containing the data for 2017 and 2018. Scope of the project is to get the percentage change of the prices of each company. Percentage change is calculated based on the price at the start of the year and price at the end of the year. By doing calculating the yearly returns of all the companies, will better help in taking informed decision rather than just baised on the bais.

### Purpose
Stock Performance of each ticker symbol for selected year. At the same time refactor the initial VBA Script which was taking more time somewhere between 40 to 50 seconds per year. Refactoring is needed so it helps to run stock analysis for many years as new data comes in.

## Stock Performance and Measuring the Refacored VB Macro
<b> Below are some of the challenges had gone through: </b> <br/>
For calculating the stock performance of the prices or returns, we need to first get all the stocks symbols , sort the data by ticker symbol, date in ascending order and loop the data for each ticker symbol then calculate the starting price and ending prices of the ticker within a year along with total volumes.

### Stock Performance
<b> 2017 Stock Analysis</b> <br/>
![2017_StockPerformance](/resources/2017_StockPerformance.png) <br/$
In 2017 except TERP company all other companies have performed well with DQ being the highest returns of 199.4% followed by SEDG with returns of 184%
<b> 2018 Stock Analysis</b> <br/>
![2018_StockPerformance](/resources/2018_StockPerformance.png)<br/>
In 2018, only couple of stocks performed, rest all tanked. In 2017 though DQ has performed extraordinarily with 199% returns, but in 2018, it was down by 62%, based on this comparision, would like to rethink on investing in DQ this year. By comparing the returns of 2017 and 2018, ENPH and RUN are the two stocks which has given consistent returns, will consider these two stocks while investing.

### Refactor Measurements
![2017_BeforeRefactoring](/resources/2017_BeforeRefactoring.png)  ![VBA_Challenge_2017](/resources/VBA_Challenge_2017.png) </br>
Based on the above results, the image on the left is taken when stock analysis is performed on the inital VB code for 2017 year, the time it has taken more than 36 seconds whereas the image on the right is taken after refactoring the code for 2017, duration for executing the analysis is less than 0.2 seconds. Similarly the same has been observed for 2018, before refactoring the code, its taking around 41 seconds to complete the process, where as after refactoring the code, its less than 0.4secs, below are the screenshots of execution times for 2018 before and after refactoring the VBA Scripts. <br/>
![2018_BeforeRefactoring](/resources/2018_BeforeRefactoring.png)  ![VBA_Challenge_2018](/resources/VBA_Challenge_2018.png) <br/>
Also made sure that the data is same before and after the refactoring the code.
![2017_StockPerformance](/resources/2018_StockPerformance.png) ![2018_StockPerformance](/resources/2018_StockPerformance.png) <br/>

## Results
- <b>Advantages and Disadvantages of Refactoring</b> <br/>
   * Advantages: <br/>
      * Less time in executing the application or getting the results back in less time.
      * Removing the redundant code helps the code to look cleaner and lean
      * Commenting along with refactoring will help other developers to understand why this has done in certain way.
      * Can be scaled to many users when rafactoring the enterprise applications.
    * Disadvantages: <br/>
      * Refactoring takes more time to improvize the performance of the application than it initially took to build/develop.
      * When developing the projects with tight deadlines, might not happen as expected.
      * Testing cycle has to be performed once again, once the refactoring is done.
- <b>Pros and Cons apply to refactoring the original VBA Script</b> <br/>   
    * Pros: <br/>
      * Took less than 0.2 seconds to perform the stock analysis for the year selected.
      * Was able to loop through entire data one time, instead of 12 times.
      * Because the time to perform the stock analysis was , able to run for multiple sceanrios without loosing the focus.
      
    * Cons:  <br/>
      * It took more time to refactor the code than it initially took to build the application.
      * Compare the results after refactoring the code with the original.
      * Make sure the code without any issues as it was doing before.
      
      
      
  



