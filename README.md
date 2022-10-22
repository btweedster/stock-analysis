# VBA Refactoring Analysis

## Purpose
The purpose of this project is to compare the differences in run times for a VBA script and a refactored version of that script.

## AllStocksAnalysis() Code

### Description
`AllStocksAnalysis()` looks at data contained in two worksheets in [VBA_Challenge.xlsm](VBA_Challenge.xlsm). There is data for 12 "green stocks" for two years: 2017 and 2018. When the script is run it will return for each stock the total Daily Volume traded for the year and the yearly return.

### Code


## AllstocksAnalysisRefactored() Code

### Refactoring

### Expectations of Refactoring

## Results

### 2017 Data
Original - 287ms

Refactored - 44ms

### 2018 Data
Original - 285ms

Refactored - 43ms

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

##### --Disadvantages--
Disadvantages to this code (as well as to the original code) are that the stock tickers must be known, and they must be sequential.