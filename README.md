# Stock Analysis with Excel VBA

## Overview of Project

### Purpose

This sheet is created to help Steve and his parents to analyze over a certain stocks to determine if they are worth investing in the year of 2017 and 2018. The set of codes, which were originally created for ticker DQ, can be used to the rest of the intended stocks. However, that set of codes was refactored to save memory as well as to make the computer to run smoother, better and faster.

### The Data

Two sets of data are presented and have the exact same format. Each set includes the 12 stocks information with the year of 2017 and 2018. Information includes such as each stock's value, the date the stock was issued, the opening, closing and adjusted price, the highest and lowest of price; and the volume of each stock. The task is to retrieve each ticker with it's daily volumne as well as the returning. 

## Results

### Analysis

The neccessery steps before refactoring are create an outline for the code, which is the comment section. Then, the previous codes are copied from the old set of codes, which includes input box, headers, ticker array and activation of the appropriate worksheet. Below is step by step.

#### 1. The 'tickers' is set equal to zero before looping over rows
<img width="168" alt="Screen Shot 2022-04-16 at 4 33 53 PM" src="https://user-images.githubusercontent.com/102835776/163694337-1bd7639b-827f-4744-9921-9fd46c7b452a.png">


#### 2. Arrays are created for 'tickers', 'totalVolumes', 'startingPrice'and 'endingPrice'
<img width="378" alt="Screen Shot 2022-04-16 at 5 36 49 PM" src="https://user-images.githubusercontent.com/102835776/163695475-2e1272c9-905f-4cf5-a0fd-7e76af9a30c7.png">

#### 3. The 'tickers' is used to access the stock ticker index for the 'tickers', 'totalVolumes', 'startingPrice'and 'endingPrice' arrays
<img width="373" alt="Screen Shot 2022-04-16 at 5 09 08 PM" src="https://user-images.githubusercontent.com/102835776/163694960-343feda8-a5f6-42fa-81dd-658271a4c830.png">


#### 4. The Script loops through stock data, reading and storing all of the following values from each row: 'tickers', 'totalVolumes', 'startingPrice'and 'endingPrice'
<img width="456" alt="Screen Shot 2022-04-16 at 5 10 38 PM" src="https://user-images.githubusercontent.com/102835776/163694991-efafb546-1ead-42d1-acdf-71246037d858.png">


#### 5. Code for formatting the cells in the spreadsheet is working
<img width="355" alt="Screen Shot 2022-04-16 at 5 12 29 PM 1" src="https://user-images.githubusercontent.com/102835776/163695013-173abfd2-c082-4aa5-8323-bbb318b19edb.png">


#### 6. There are comments to explain the purpose of the code
<img width="470" alt="Screen Shot 2022-04-16 at 5 13 34 PM" src="https://user-images.githubusercontent.com/102835776/163695037-af89f08e-4157-4bd9-8709-8318d5b3a912.png">


#### 7. The outputs for the 2017 and 2018 stock analysis in the 'VBA_Challenge.xlsm' workbook match the outputs from the AllStockAnalysis in the module
<img width="570" alt="Screen Shot 2022-04-16 at 5 14 54 PM" src="https://user-images.githubusercontent.com/102835776/163695062-5043a155-64a2-4340-b4fd-0c1e5f3f951d.png">
<img width="568" alt="Screen Shot 2022-04-16 at 5 15 53 PM" src="https://user-images.githubusercontent.com/102835776/163695082-96711d36-709c-49c4-a4c5-69e1826de33a.png">


#### 8. The pop-up messages showing the elapsed run time for the script are saved as 'VBA_Challenge_2017.png' and 'VBA_Challenge_2017.png'
<img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/102835776/163695089-2e5ca9ab-2caf-419e-b171-edb49035c12a.png">
<img width="247" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/102835776/163695093-0119005e-78d7-4a70-8d97-5f76e6e3ed8c.png">

## Sumary

### Pros
- Cleaner
- More organized
- Software improvement
- Software debugging
- Faster programing
- Benefits users who view the project since it's more straightforward
- Decrease in macro run time

### Cons
- Can affect the testing outcomes
- Not having proper test cases for the existing codes when the application is too large
