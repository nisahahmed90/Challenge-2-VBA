# Green-Energy Stocks Analysis
## Overview of the Project
### Background
In this project, I am performing an analysis of green-energy stock for Steve from years 2017 and 2018. The analysis is aimed at the yearly return and total daily volume. 

For this data analysis I have used Microsoft Visual Basic (VBA), including conditional statements, for loops, static and conditional formatting and code refactoring in order to improve its clarity and efficiency. 

VBA is used to automate the tedious processes, improve the efficiency and uniformity of the analysis output, reduce the chances of accidents and error and to write a code that can be used for similar future projects. We can come back to the “old code” and rewrite to make it work better. In ![VBA_Challenge] (https://github.com/nisahahmed90/Challenge-2-VBA/blob/main/VBA_Challenge.xlsm) file are two vbs modules that contain VBA code before refactoring (Module 1) and after refactoring (Module 3). 

### Purpose
The purpose of this analysis is to help Steve analyze the green-energy stock market for his parents. They are interested in investing in DAQO stocks (Ticker: DQ), a company that makes silicon wafers for solar panels. Before investing their money, Steve wants to run some analyses and check DQ stocks performance over the years in comparison to other green-energy stocks. Results will help him determine if DQ stocks are worth investing his parents’ money.

## Results
### Analysis of green stocks for 2017 and 2018

The tables below display the analysis for a 12 different green-energy stocks. The tables contains three groups of data:
	▪	Ticker name
	▪	Total Daily Volume
	▪	Percentage of a yearly return

![](https://github.com/nisahahmed90/Challenge-2-VBA/blob/main/All%20Stock%202017.png)
![](https://github.com/nisahahmed90/Challenge-2-VBA/blob/main/All%20Stocks%202018.png)

#### Return
Green-energy stocks in 2017 had a high ratio of positive yearly returns (only one green-energy stock (TERP)) had a negative yearly return. Analysis from 2018 showed a different picture altogether. The majority of stocks had negative returns. The DQ stock had almost 200% yearly return in 2017, but in 2018 the stock dropped and finished the year with negative 63%.

These results indicate a risky investment. The stock trend is not stable and might not be worth investing all the money in DQ stocks.

#### Daily Volume
In general, a high volume of daily trading is an indicator of a stable stock, with a lot of interest and activity.

DQ stocks in 2017 had low volume and high yearly return. However, the situation of DQ stocks in 2018 has changed completely. Stocks closed its year with negative 63%. Trading volume was higher, yet didn’t result in a positive outcome. The results of this analysis confirmed a risky investment in DQ stocks.

### Refactoring the Code
Both scripts “AllStockAnalysis” and “AllStockAnalysisRefactored” have the same output. 

To make my code more efficient, I created 3 new arrays: -tickerVolumes(12) to hold volume -tickerStartingPrices(12) to hold starting price -tickerEndingPrices(12) to hold ending price.
The above 3 arrays store performance data for each stock when a for loop runs analysis on them. Matching the 3 performance arrays with the ticker array is done by using a variable called the tickerIndex. After creating these arrays I used Nested For Loops and variables to loop through the data and complete the analysis.

![](https://github.com/nisahahmed90/Challenge-2-VBA/blob/main/Screen%20Shot%202022-03-20%20at%2011.47.22%20PM.png)
![](https://github.com/nisahahmed90/Challenge-2-VBA/blob/main/Screen%20Shot%202022-03-20%20at%2011.48.19%20PM.png)


## Summary
### Advantages and Disadvantages of refactoring code?

The purpose and the advantages of refactoring code are to improve code:
	•	efficiency - code is taking fewer steps, therefore taking up less computer memory and taking-up less time to execute the code,
	•	readability - code is easier to understand, it’s cleaner as a result of improved logic of the code,
	•	functionality - fixing any bugs that might have been overlooked in the original code.
On the other hand, the disadvantages of refactoring code can be:
	•	frustrating and time-consuming - we might not be aware of the purpose of the code and its functionality. Especially when the code is not well commented and we could spend a lot of time figuring out what specific lines or blocks of code are supposed to do. That's why the good documentation and commenting the code is very important.
	•	less efficient - by refactoring the code, we could end up with a less efficient script.

### How do these pros and cons apply to refactoring the original VBA Script

There are pros and cons to both sides. Refactoring VBA scripts, especially for beginners (like myself -- at this point), requires quite a bit of effort. At some point was a bit frustrating and confusing, since the understanding of the basics weren’t under the belt yet. Yet, on the other hand, was extremely rewarding and fulfilling. This technique added up another
level of coding -- that is -- deepened the understanding of the logic of the code. Hard work was paid off with new knowledge and understanding of the complex structure of the code. Moreover, I was able to see improved efficiency immediately. By rewriting the script we were able to avoid nested loops, so the code wasn’t switching back and forth between 
the worksheets (that can be quite process-intensive). The new code ran much faster, 5-times faster and because of this, we could reuse the new code on a much larger data set. Another bigger improvement of the code was accessing the arrays with a single variable tickerIndex. In this case, code stored all elements in arrays before switching to another
worksheet.

