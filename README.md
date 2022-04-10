# All Stocks Analysis 2017 and 2018

## Project Overview

Steve is manageing his parents stock portfolio. They are passsionate about green energy and want to invest in DAQO New Energy Corporation (DQ). They have not done any research on DQ and simply chose it because they met many years ago at Dairy Queen. Another DQ. Steve's goal is to do a thorough analysis of DQ as well as other green energy stocks so that he can diversify his parents portfolio. 

Steve likes the work I have done so far to analyze a few dozen stocks but would like to see how the analysis of thousands of stocks would work. Utilizing refactoring code, my goal is to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Results

In comparing DQ for years 2017 and 2018, 2017 had a better return at a positive 199.4%. While 2018 fared worse at a lose of return of -62.6%. The same could be said for all other stocks. 2017 was definitely a better year for all stocks except for TERP that had a negative return of -7.2%.

For 2018 all stocks performed at loss of return, except for ENPH and RUN who both had a positve return. 

![All Stocks 2017](https://user-images.githubusercontent.com/100816778/158927782-6498ec71-20ea-4c9f-b025-1991ff6bb3d3.png)
![All Stocks 2018](https://user-images.githubusercontent.com/100816778/158927792-97c94d16-6f34-40d3-8581-c2405dd00fec.png)

In order to analyze many stocks, refactoring code was done. A input box was created on the All Stocks Analysis worksheet that asks the question "What year would you like to run the analysis on?" This box appears when the "Run Analysis for All Stocks" button is pushed. The following code was used to create the input box:

yearValue = InputBox (What year would you like to run the analysis on?")

The user then enters the year that they want ran, and the information is returned for that year. By adding a message box using the following code:

MsgBox "This code ran in " & (endTime-startTime) & "seconds for the year" & (yearVAlue)

The run time for the code is returned. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/100816778/158929595-22cfa97f-0667-4c98-b4da-c500773e7da6.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/100816778/158929606-68af48ad-9be3-41fe-b141-791258404cc5.PNG)

## Summary
1. What are the advantages or disadvantages of refactoring code?

The advantage of refactoing code is to run the code more efficiently when you have a large amount of data to loop through. When there is a large amount of data to loop through this could slow up your computer or even freeze it up.

2. How do these pros and cons apply to refactoring the original VBA script?

In refractoring the original VBA script, we are able to see how much time it took to return what was asked. 
