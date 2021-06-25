# VBA of Wall Street - Stock Analysis

## Overview of Project

### Purpose

Steve recently graduated with a finance degree and wants to help his parents invest in green energy stocks. He was presented with a workbook that could analyze 12 stocks. For him to make a better recommendation to his parents on which stocks to invest in, he wants to increase the volume of his dataset. The existing VBA script might not perform as expected if used for an analysis with a large dataset. In order to help Steve, the VBA script must be refactored.

The purpose of this project was to refactor the existing solution code and to see if refactoring made the script more efficient or not. No new functionality was added. Script performance was measured by their runtimes and was compared with the runtimes of the initial script to do the same analysis on the 12 stocks. 

*Files:*
- [Stock Analysis Worksheet](VBA_Challenge.xlsm)
- [Existing Solution Code](https://github.com/samanthajpv/Stock-Analysis/blob/7965c5472d6b09704e8b02932786e1943a8f7a00/Resources/Additional/Existing%20Solution%20Code.vbs)
- [Refactored Script](VBA_Challenge.vbs)



## Results

### Stock Performance

| Stock Analysis 2017  | Stock Analysis 2018 |
| ------------- | ------------- |
| <img src="https://github.com/samanthajpv/Stock-Analysis/blob/7965c5472d6b09704e8b02932786e1943a8f7a00/Resources/Additional/Stock%20Analysis%202017.png" width="250" height="250">|<img src="https://github.com/samanthajpv/Stock-Analysis/blob/7965c5472d6b09704e8b02932786e1943a8f7a00/Resources/Additional/Stock%20Analysis%202018.png" width="250" height="250">|

The *Ticker* column represents the stock symbol while the *Total Daily Volume* was calculated by adding the volume of traded stocks per day throughout the year. On the other hand, _**Return**_ is comprised of the closing stock prices and was calculated by the line of code below: 

```Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1```

Looking at the returns, 2017 was a good year for green energy stocks with majority having positive returns. Although, 2018 did not do well. The cells were formatted to be red for a negative return and green for a positive return. There is only one negative return in 2017 while 10 out of the 12 stocks in 2018 have negative returns. That is 83% of the dataset. Based on this analysis, it is safe to say that ENPH and RUN stocks are promising stocks to invest in since the two have positive returns for both years. But of course, there are other factors to consider as well. It is also important to have a diversified stock porfolio given that the analysis was only for 12 green energy stocks.

### VBA Script - Existing Solution Code

| Existing Script Runtime 2017  | Existing Script Runtime 2018 |
| ------------- | ------------- |
| <img src="https://github.com/samanthajpv/Stock-Analysis/blob/ac17686c161d9da5a2e187b2c2385824c954db2c/Resources/Additional/Initial%20Code%202017.png" width="320" height="130">|<img src="https://github.com/samanthajpv/Stock-Analysis/blob/ac17686c161d9da5a2e187b2c2385824c954db2c/Resources/Additional/Initial%20Code%202018.png" width="320" height="130">|

The user is prompted to input the desired year for analysis (2017 or 2018) through clicking a macro button. Once entered, the code will run and return a message box with the run time of the script. For the existing solution code, the runtime for analyzing the 2017 data is 0.5859 seconds while 0.6016 seconds for 2018.

<img src="https://github.com/samanthajpv/Stock-Analysis/blob/35487e6618347d82e70d905d616295d9ea5d68b8/Resources/Additional/Existing%20Solution%20Code%20-%20Nested%20For%20Loop.png" width="370" height="400">

The existing solution code has a nested for loop. It starts from looping through the tickers (stock symbols) to looping through rows of data with conditional code. The inner for loop sums up the volume if the cell contains the ticker that is currently being checked and stores the starting and ending price. The outer for loop continues up to displaying the results of the analysis in a different worksheet. 

### VBA Script - Refactored

| Refactored Script Runtime 2017  | Refactored Script Runtime 2018 |
| ------------- | ------------- |
| <img src="https://github.com/samanthajpv/Stock-Analysis/blob/ac17686c161d9da5a2e187b2c2385824c954db2c/Resources/VBA_Challenge_2017.png" width="320" height="130">|<img src="https://github.com/samanthajpv/Stock-Analysis/blob/ac17686c161d9da5a2e187b2c2385824c954db2c/Resources/VBA_Challenge_2018.png" width="320" height="130">|

The same data was used and the same output was created with the refactored script to ensure that performance can be compared against that of the existing script. With the refactored code, the runtime for analyzing 2017 and 2018 data went down to 0.1172s and 0.125s respectively.

```
 '1a) Create a ticker Index
    'setting tickerIndex to zero before it iterates through the rows
    tickerIndex = 0

    '1b) Create three output arrays
    'declaring arrays to store the values for the analysis
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
```
The first modification in the script was the utilization of indexing and output arrays. A ```tickerIndex``` was created and arrays for the volume and prices. Arrays use in-memory processing which makes it a better performer compared to looping through cells *(Sagmon, 2020)*. The index is used to call each element in the array which can be seen in the code below.

<img src="https://github.com/samanthajpv/Stock-Analysis/blob/35487e6618347d82e70d905d616295d9ea5d68b8/Resources/Additional/Refactored%20Script%20-%20For%20Loops.png" width="550" height="400">

The second modifcation was the breaking of the nested loop. The refactored code now has three *For Loops* instead of one big nested loop. This was possible in this case because of the use of the output arrays. Also, the conditional statement for the volume was replaced with a single formula:

``` tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value ``` 

*(Formula from https://courses.bootcampspot.com/courses/626/assignments/13362?module_item_id=211669)*

### VBA Script - Comparison

| Year | Existing Script Runtime | Refactored Script Runtime | Runtime % Difference |
| ------------- | :-----------: | :-----------: | :-----------: |
| 2017 | 0.5859375s | 0.1171875s | 80.0% |
| 2018 | 0.6015625s | 0.1250s | 79.2% |

Runtimes for analyzing the dataset decreased by 80% for 2017 and 79.2% for 2018. Refactoring the code was indeed successful. The refactored code is roughly 5 times faster than the initial solution code and this is a huge improvement in script performance. For larger datasets, it is recommended to use the refactored script.


## Summary
 
The purpose of code refactoring is to clean the code by making it readable, increase its performance, and make it extendable *("Code Refactoring: Meaning, Purpose, Benefits", n.d.)*. It helps developers add new features easily because of the reusable code which can save a lot of time. A disadvantage of code refactoring is that is does take time. Depending on how easy it is to understand the existing script, detect the bugs, and assess its maintainability, some scripts may not be worth editing at all. 

For this stock analysis, it can be concluded that refactoring the original VBA script made it more efficient in terms of runtime. The script was very straightforward making it understandable. Also, the goal to be able to expand the dataset and do the same analysis, the refactored code will perform better than the orginal code.


## References

(1) Trilogy Education Services. (2021, June). *Module 1 Challenge*. https://courses.bootcampspot.com/courses/626/assignments/13309?module_item_id=211540

(2) Sagmon, M. (2020, November 12). Excel VBA: Arrays. *MorSagmon*. https://www.morsagmon.com/blog/Excel-VBA-Arrays

(3) *Code Refactoring: Meaning, Purpose, Benefits*. (n.d.). Sumatosoft. Retrieved June 24, 2021, from https://sumatosoft.com/blog-post/code-refactoring-meaning-purpose-benefits
