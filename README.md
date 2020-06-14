Challenge Assignment - Module 2
---
Initially, the code was written to assist Steve with his analysis of green stocks. The data provided had known number of stocks, therefore, when initiating a data array, we enlisted all known stocks within it. 
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

The search was constructed with a nested `For` loop. The outer `For` loop was going through each stock, passing it into inner `For` loop and testing it through all the rows of a dataset, assigning appropriate variables back into the single dimension array if conditions were satisfied.

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

In terms on efficiency and number of operations, it is proportional to the **"number of stocks" x "number of rows"** in a data set. In this case, it was **"12" x "3012" = 36,156 operations.**

Since Steve wanted to do a little more research for his parents and was gathering data for the entire stock market over the last few years, we needed to refactor the intial code. Since we assumed that there is an unknown number of stocks, the previous code which had a defined array and looped over a number of stocks is no longer valid. Therefore, first of all we had to initiate a dynamic array `As Variant` type because the final number of stocks is not known, although we do know that ultimately we need to include four variables/dimensions into it (stock, opening price, closing price, volumes). We also had to introduce tickerIndex variable in order to later use it within our `For` loop for switching between stock tickers.
```
   Dim ticker() As Variant
   tickerIndex = 0
```
Once initilized, we had to re-dimension the array to create placeholders for the 4 variables that we needed to perform the analysis on. 
```
ReDim ticker(4, tickerIndex)
```

we had to construct an algorythms which only looped over the number of rows once, performing necessary tests, incrementing the number of stocks as necessary and assigning the appropriate values into the array. In terms on efficiency and number of operations,
this new algorythm is proportional only to the number of rows in the data set, in this case of test data set 3012 only.


