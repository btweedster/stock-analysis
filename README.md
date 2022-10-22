# VBA Refactoring Analysis

## Purpose
The purpose of this project is to compare the differences in run times for a VBA script and a refactored version of that script.

## AllStocksAnalysis() Code

### Description
`AllStocksAnalysis()` looks at data contained in two worksheets in [VBA_Challenge.xlsm](VBA_Challenge.xlsm). There is data for 12 "green stocks" for two years: 2017 and 2018. When the script is run it will list for each stock the total Daily Volume traded for the year and the yearly return expressed as a percentage.

### Code

## AllstocksAnalysisRefactored() Code

### Refactoring

### Expectations of Refactoring

## Results

### 2017 Data
Original - ~287ms

Refactored - ~55ms

### 2018 Data
Original - ~285ms

Refactored - ~55ms

## Advantages and Disadvantages of Refactoring

### General Considerations

#### --Advantages--

#### --Disadvantages--

### Considerations of Refactoring above VBA Code

#### --Original Code--

##### --Advantages--

##### --Disadvantages--
Disadvantages to both the original and refactored code to be discussed below.

#### --Refactored Code--

##### --Advantages--
Per the above analysis, the most obvious advantage of the refactored code is the improved speed. ~300ms to ~50ms may not seem like a signifcant practical improvement - and for 3012 rows of data, it's not. However, we should consider that this script could be run on data with an arbitrary number of rows. And as the number of rows increases, we would expect to see the runtime for the original code to grow considerably more quickly than the runtime of the refactored code.

##### --Disadvantages--
Disadvantages to this code (as well as to the original code) are that the stock tickers must be known, and they must be sequential.