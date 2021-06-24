# VBA of Wall Street - Stock Analysis

## Overview of Project

### Purpose

Steve recently graduated with a finance degree and wants to help his parents invest in green energy stocks. He was presented with a workbook that could analyze 12 stocks. For him to make a better recommendation to his parents on which stocks to invest in, he wants to increase the volume of his dataset. The existing VBA script might not perform as expected if used for an analysis with a large dataset. In order to help Steve, the VBA script must be refactored.

The purpose of this project was to refactor the existing solution code to see if it will make the script more efficient. Script performance was measured by their run times and was compared with the run times of the initial script to do the same analysis on the 12 stocks. 

*Files:*
- [Stock Analysis Worksheet](VBA_Challenge.xlsm)
- [Existing Solution Code](https://github.com/samanthajpv/Stock-Analysis/blob/7965c5472d6b09704e8b02932786e1943a8f7a00/Resources/Additional/Existing%20Solution%20Code.vbs)
- [Refactored Script](VBA_Challenge.vbs)

## Results

### Stock Performance

| Stock Analysis 2017  | Stock Analysis 2018 |
| ------------- | ------------- |
| <img src="https://github.com/samanthajpv/Stock-Analysis/blob/7965c5472d6b09704e8b02932786e1943a8f7a00/Resources/Additional/Stock%20Analysis%202017.png" width="250" height="250">|<img src="https://github.com/samanthajpv/Stock-Analysis/blob/7965c5472d6b09704e8b02932786e1943a8f7a00/Resources/Additional/Stock%20Analysis%202018.png" width="250" height="250">|

The *Ticker* column represents the stock symbol while the *Total Daily Volume* was calculated by adding the volume of traded stocks per day throughout the year. On the other hand, *Return* was calculated by dividing the ending closing price of the stock for the year by the starting closing price of the year and subtracting 1. 

Looking at the returns, 2017 was a good year for green stocks with majority having positive returns. Although, 2018 did not do well. The cells were formatted to be red for a negative return and green for a positive return. There is only one negative return in 2017 while 10 out of the 12 stocks in 2018 have negative returns. That is 83% of the dataset. Based on this analysis, it is safe to say that ENPH and RUN stocks are promising stocks to invest in since the two have positive returns for both years. But of course, there are other factors to consider as well. It is also important to have a diversified stock porfolio given that the analysis was only for 12 green energy stocks.

### VBA Script - Existing Solution Code
(insert pics run times for existing code) + examples of code

The user is prompted to input the desired year for analysis through clicking a macro button. Once entered, the code will run and return a message box with the run time of the script. For the existing solution code, the run time for analyzing the 2017 data is 0.5859 seconds while 0.6016 for 2018.

The existing solution code has a nested for loop from looping through the tickers to looping through rows with conditional code up to displaying the results of the analysis. This creates a big

### VBA Script - Refactored
(insert pics run times for refactored code) + examples of code


### VBA Script - Comparison
(insert table for run time difference)


## Summary
In a summary statement, address the following questions.
    1. What are the advantages or disadvantages of refactoring code?
    2. How do these pros and cons apply to refactoring the original VBA script?


## References
