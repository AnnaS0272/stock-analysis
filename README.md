Module 2 - Challenge Assignment
---
Initially, the code was written to assist Steve with his analysis of green stocks. The data provided had a known number of stocks, therefore, when initiating a data array, we enlisted all known stocks within it. 
```
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

The search was constructed with nested `For` loops. The outer `For` loop was going through each ticker, passing it into the inner `For` loop which was testing the ticker through all the rows of the dataset, calculating appropriate variables and returning outputs into the spreadsheet after the inner `For` loop, providing conditions were satisfied.

```
For I = 0 To 11
        ticker = tickers(I)
        totalVolume = 0
        Worksheets(yearValue).Activate
        
        For J = 2 To RowCount
            If Cells(J, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(J, 8).Value
            End If
            If Cells(J - 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then
                startingPrice = Cells(J, 6).Value
            End If
            If Cells(J + 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then
                endingPrice = Cells(J, 6).Value
            End If
        Next J
 Next I
 ```

In terms of efficiency and number of operations, it is proportional to the **"number of stocks" x "number of rows"** in a data set. In this case, it was proportional to **"12" x "3012" = 36,156 operations.**

## Refactoring the code

Since Steve wanted to do a little more research for his parents and was gathering data for the entire stock market over the last few years, we needed to refactor the intial code. Since we assumed that there is now an unknown number of tickers, the previous code which had assumed a number of known tickers is no longer valid. Therefore, first of all we had to initiate a dynamic 2-dimensional array `As Variant` type because the final number of tickers was not known. We did know that ultimately we needed to include four variables into the array (stock, opening price, closing price, volumes). We also had to introduce tickerIndex variable in order to later use it within our `For` loop for switching between stock tickers and outputting to the spreadsheet afterwards.
```
   Dim ticker() As Variant
   tickerIndex = 0
```
Once initialized, I had to re-dimension the array to create placeholders for the 4 variables that I needed to perform the analysis on. 
```
ReDim ticker(4, tickerIndex)
```
Rather than having four independant arrays we refactored our code with a single 2-dimensional array with the following indexing.
```
'ticker(1,tickerIndex) = ticker name
'ticker(2,tickerIndex) = opening price
'ticker(3,tickerIndex) = closing price
'ticker(4,tickerIndex) = volume
  ```
I also had to create some starting point in our `For` loop analysis, i.e., I had to start with an initial tickername and opening price. I assumed the first tickername we are running is in cell A2. the reson we needed to segment the opening price for the first ticker out of the loop, so it doesn't get overwritten as we loop through, which would produce incorrect financial calculations.
```
ticker(1, tickerIndex) = Cells(2, 1).Value
ticker(2, tickerIndex) = Cells(2, 3).Value
```
I then constructed an algorithm which only looped over the number of rows once, performing necessary tests, incrementing the number of stocks via `tickerIndex` as necessary and assigning the appropriate values into the array. In terms of efficiency and number of operations, this new algorithm is proportional only to the number of rows in the data set, in this case of test data, the set is **only 3012 rows.** Compared to the previous version of the code, this refactored version runs 12 times faster.

```
For I = 2 To RowCount
        If ticker(1, tickerIndex) = Cells(I, 1).Value Then
            ticker(3, tickerIndex) = Cells(I, 6).Value 
            ticker(4, tickerIndex) = ticker(4, tickerIndex) + Cells(I, 8).Value 
        Else 
            tickerIndex = tickerIndex + 1 
            ReDim Preserve ticker(4, tickerIndex) 
            ticker(1, tickerIndex) = Cells(I, 1).Value 
            ticker(2, tickerIndex) = Cells(I, 3).Value 
            ticker(3, tickerIndex) = Cells(I, 6).Value 
            ticker(4, tickerIndex) = ticker(4, tickerIndex) + Cells(I, 8).Value 
        End If
    Next I
 ```

The last thing to note, is that I had to introduce a redimensioning within the `Else` part of the loop, when the new ticker is identified via a tickerIndex logic. The redimensioning was important as it would create the new entry. It is needed to `ReDim` every time there is a new ticker because it's a dynamic array an we don't know how many tickers there are or will be.

Now Steve is all equipped and ready to go with his All Stock Analysis.




