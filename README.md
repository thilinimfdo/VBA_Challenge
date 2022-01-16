ANALYSIS OF GREEN ENERGY INVESTMENT OPPORTUNITES 

1.Overview of the Project
Steve’s Parents are interest on investing in green energy and has requested Steve to conduct research on alternate investing opportunities in Green Energy. However, Steve decided to invest his parent’s money in Daqo Energy Corporation (a Company makes Silicon Wafers and Solar Panels) while diversifying the funds in alternate green energy stocks.

 This Report will Analyses 
•	The stock performance of Daqo Energy Corporation (DQ) and the other alternate green energy investments 
•	The performance of the VBA code written with its Refactored VBA Code 

2. Results

2.1 Stock Performances between 2017 and 2018

Looking at the “Total Daily Volume” performance, it shows that certain organization’s (Ticker’s) transaction volume have increased and some’s not. The most significant rise in volume transaction can be seen in Ticker DQ (201% increment in volume transaction in 2018 compared to 2018) and in Ticker ENPH (174% increment in volume transaction in 2018 compared to 2017) 

However, in observing the return, DQ has made a negative Yearly Return performance in compared with its 2017 performance (2017 Yearly Return : 199.4% & 2018 Yearly Return : -62.6%). This raise an alarm in investing in DQ stocks. (Refer https://github.com/thilinimfdo/VBA_Challenge/blob/main/Resources/VBA_Challange_2017.png & https://github.com/thilinimfdo/VBA_Challenge/blob/main/Resources/VBA_Challage_2018.png)

•	Why the Stock transaction Volume has increase?

•	But not the Yearly Return?

This can be assumed that the organization financial performance may not healthy and as a result the current investors are trying to sell their stocks to move out from DQ. And as the demand for DQ stock does not Rise as the Volume of Stocks put the market to sell may drop the Stock prices of the company. This may not a favourable alarming to invest in DQ. 

However, the Yearly Return value of ENPH indicates a positive figure though there is a drop of the percentage. (129.5% Yearly Return in 2017 and 81.9% yearly Return in 2018)
  
Looking at the performance of DQ compared with ENPH, it indicates as ENPH as a better alternate of investment opportunity compared to DQ. 

However, it is recommended Steve to deep dive in DQ other performance factors such as their investments made currently to gain higher return to the future. For example, if the organization is not performing well because they have heavily invested now to gain a higher return in the future, there is a higher chance of the return value to be increase significantly when the said investments start to yield. OR This performance may could be due to mismanagement practices of the organization. As there is a drop in Yearly Return of many organizations in the datasheet, it also could assume that the economy of the country may not had been favourable for the green energy industry and had been a challenging year.

Further, it could recommend to look at last five years performance to understand a trend of “Total Daily Volume” and “Yearly Return” and Steve to deep dive in other aspects of organization performance before deciding the stocks that he would invest his parent’s money.

2.2 Execution Times of the Original Script and the Refactored Script

The Time Taken to run the Original Code for 2017 is 1.037109 Seconds and for2018 is 1.044922 Seconds (https://github.com/thilinimfdo/VBA_Challenge/blob/main/Resources/Time_Taken_2017_Analysis_Original_Code.png & https://github.com/thilinimfdo/VBA_Challenge/blob/main/Resources/Time_Taken_2018_Analysis_Original_Code.png)

The Time taken to run the Refactored Code for 2017 is 0.1894531 Seconds and for 2018 is 0.1992188 Seconds (https://github.com/thilinimfdo/VBA_Challenge/blob/main/Resources/Time_Taken_2017_Analysis_Refactored_Code.png & https://github.com/thilinimfdo/VBA_Challenge/blob/main/Resources/Time_Taken_2018_Analysis_Refactored_Code.png)

Therefore, the refactored code run has dropped by 0.8476559 Seconds for 2017 and 0.8457032 Seconds for 2018 Analysis. (0.85 Seconds Drop for both Analysis or nearly 90% of the original execution time)

The Main Change made to the Refactored Code is 
•	In Original the arrays were created for Tickers and in the Refactored code the array was created for both tickers and results.
In the original code, after each iteration, the code changes active worksheet and prints out the results (Volume, Starting Price, End Price) since the code holds the results in a single value variable. In the refactored code, the result is stored in an array enabling us to store the results without needing to print each iteration. Hence, we can finish going over the data without printing and worry about printing the results and changing the worksheet after the loop. In my opinion, this has resulted in a significant performance increment – reduced running time.
	
3. Summary
3.1 Advantages and Disadvantages of Refactoring Code (Detail Statement)
•	Advantages of Refactoring Code
o	As per the discussion in 2.2, the efficiency of the running code can increase Significantly
o	The Macro will not be long and easy to view the whole subroutine 

•	Disadvantage of Refactoring Code
o	As the Coding is not long enough (not detailed), there is a high possibility of getting Syntax Errors and difficult in correcting such Syntax errors
o	Apart from the coding writer, the readability of the macro for others is less. Therefore, it would be challenging for another person to change the code in the future.

3.2 Advantages and Disadvantages of the Original and Refactored VBA Script
The only disadvantage is the during the printing the results, we are activating the “All Stocks Analysis” each iteration in the loop which is unnecessary. We could just activate the sheet once before the final loop.
