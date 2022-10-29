# VBA Refactoring Analysis

## Purpose
The purpose of this project is to compare the differences in run times for a VBA script and a refactored version of that script.

## AllStocksAnalysis()

### Description
`AllStocksAnalysis()` looks at data contained in two worksheets in [VBA_Challenge.xlsm](VBA_Challenge.xlsm). There is data for 12 "green stocks" for two years: 2017 and 2018. When the script is run it will list for each stock the total Daily Volume traded for the year and the yearly return expressed as a percentage. Formatting is added to to the yearly return to be demonstrate a stock's performance. 

### AllStocksAnalysis() Code
Generally, the scipt runs by having an array of all the stock tickers that we wish to analyze. We loop over the ticker array and within that, we loop over all of the stock data and extract the information that we need on that stock.

The following is the original code. The whole script will not be listed here, only those sections that are of primary interest to the refactoring. The lines enclosing each `for` loop are added here for emphasis and clarity.

```
'4) Loop through tickers
 ---For i = 0 To 11
|   
|       ticker = tickers(i)
|       
|       totalVolume = 0
|       
|       Worksheets(yearValue).Activate
|   
|       '5) loop through rows in the data
|    ---For j = 2 To RowCount
|   |       
|   |       '5a) Get total volume for current ticker
|   |       If Cells(j, 1).Value = ticker Then
|   |       
|   |           totalVolume = totalVolume + Cells(j, 8).Value
|   |       
|   |       End If
|   |       
|   |       '5b) get starting price for current ticker
|   |       If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
|   |       
|   |           startingPrice = Cells(j, 6).Value
|   |       
|   |       End If
|   |       
|   |       '5c) get ending price for current ticker
|   |       If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
|   |       
|   |           endingPrice = Cells(j, 6).Value
|   |       
|   |       End If
|   |   
|    ---Next j
|       
|       'Activate the output worksheet
|       Worksheets("All Stocks Analysis").Activate
|       
|       'Output the data for the current ticker
|       Cells(4 + i, 1).Value = ticker
|       Cells(4 + i, 2).Value = totalVolume
|       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
|   
 ---Next i
```

As can be seen we have a `for` loop inside of a `for` loop. This is done for each ticker. So if we were only to add tickers on which to get data (assuming we werern't already getting data on every tracker in the dataset) or if we were adding data to the dataset (without also tracking its ticker) the runtime of this script would grow linearly with the tickers or data. However, if we begin to add new data *and* the new tickers to the array for tracking, we would expect this runtime to grow *exponentially*.

## AllstocksAnalysisRefactored() Code
Wanting to avoid an exponentially increasing runtime, we have refactored the code so that an increase in data will only be a linear increase in runtime.

The following code is the refactored portion from above. Again the lines enclosing the `for` loop are added for emphasis.

```
'2b) Loop over all the rows in the spreadsheet.
 ---For i = 2 To RowCount
|   
|       '3a) Increase volume for current ticker
|       tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(i, 8).Value
|       
|       '3b) Check if the current row is the first row with the selected tickerIndex.
|       If Cells(i, 1) <> Cells(i - 1, 1) Then
|           
|           tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
|           
|       End If
|       
|       '3c) check if the current row is the last row with the selected ticker
|       'If the next row’s ticker doesn’t match, increase the tickerIndex.
|       If Cells(i, 1) <> Cells(i + 1, 1) Then
|           
|           tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
|           tickerIndex = tickerIndex + 1
|           
|       End If
|  
 ---Next i
```

### Refactoring
As can be seen, there is now only a single `for` loop, i.e. we only loop over the rows of data a single time, instead of once for every ticker. This required us to initialize several tracker variables before this refactored loop:

```
    tickerIndex = 0
    Dim tickerVolume(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```

We did this so that we can increment `tickerIndex` as soon as we reach a new ticker in the dataset. In doing so, we can access the above arrays according `tickerIndex` and arrive at the same information as the original code. 

### Expectations of Refactoring
Because we are reducing the number of loops over the dataset from 12 (the number of tickers) to 1, we expect our runtime to drop dramatically. Now this doesn't mean that expect our refactored runtime to be 1/12 of the original runtime. There are many other operations such as variable declarations and formatting that remain the same between the original and refactored code.

### Limitations
Both the original code and the refactored code assume that data for each ticker is continuous. That is, *all* data for a stock is listed before a new stock is listed. Additionally, the array `tickers` is hardcoded in the order in which the stocks appear in each data set:
```
'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
```

If each stocks data were not listed in continuous "blocks", or if those blocks were in any order other than according to `tickers`, the code would either not run, or would return incorrect information.

## Results

### 2017 Data
The runtime for the original code on the 2017 dataset is approximately 287ms. The following shows the runtime for the 2017 dataset with the refactored code is approximately 55ms. This is approximatey 20% of the original runtime.

![VBA_Challenge_2017.png](resources/VBA_Challenge_2017.png)

### 2018 Data
The runtime for the original code on the 2018 dataset is approximately 285ms. The following shows the runtime for the 2018 dataset with the refactored code is approximately 55ms. This is again approximately 20% of the original runtime.

![VBA_Challenge_2018.png](resources/VBA_Challenge_2018.png)

## Advantages and Disadvantages of Refactoring

### General Considerations of refactoring code.

#### Advantages
The main advantages of refactoring code are to improve efficiency and maintain functionality. The above refactoring is a good example of improving efficiency. Even minor improvements in efficiency can be a significant improvement if the refactored code is run often. Additionally, as other packages, modules, and software that our depends upon are refactored or changes themselves, it may be necessary to refactor our code so that the code remains functional.

#### Disadvantages
We may wish to improve code with additional or altered functionality, however when refactoring our goal is to maintain the same functionality but with improved functionality. So we must first separate the two ideas of improving functionality and improving efficiency with the same functionality. In reality, we may do both at the same time, but we must be cognizant of our goal(s) before diving in. When we decide that we are only refactoring, we are constrained to produce the same results between the original and the refactored code.

### Considerations of Refactoring above VBA Code

#### Original Code Advantages
The advantage to note concerning the original code is that it works. It gives exactly the information and formatting that we want. And while it is best practice to write efficient code, this is a good starting point. We can write something functional, and then allow the refactoring process to improve efficiency.

#### Original Code Disadvantages
The primary disadvantage of the original code is the inefficient nesting of the `for` loops. It is perfectly acceptable to nest `for` loops when warranted, however, it is not necessary here. Disadvantages to both the original and refactored code to be discussed below.

#### Refactored Code Advantages
Per the above analysis, the most obvious advantage of the refactored code is the improved performance. ~300ms to ~50ms may not seem like a signifcant, practical improvement - and for 3012 rows of data, it's not. However, we should consider that this script could be run on data with an arbitrary number of rows. And as the number of rows increases, we would expect the runtime for the original code to grow considerably more quickly than the runtime of the refactored code.

#### Refactored Code Disadvantages
Both the original code and the refactored code assume that data for each stock is continuous. That is, *all* data for a given stock is listed before the data of a different stock is listed. Additionally, the array `tickers` is hardcoded in the order in which the stocks appear in each data set:
```
    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
```

If each stocks data were not listed in continuous "blocks", or if those blocks were in any order other than according to `tickers`, the code would either not run, or would return incorrect information.