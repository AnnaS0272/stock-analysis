###Challenge Assignment - Module 2
---
Initially, the code was written to assist Steve with his analysis of green stocks, i.e., stock performance. The data provided had
known number of stocks, therefore, when initiating a data array, we enlisted all known stocks within it. The search was constructed with
a nested FOR loop. The outer FOR loop was going through each stock, passing it into inner loop and testing it through all the rows,
assigning appropriate variables back into single dimension array if conditions were satisfied. In terms on efficiency and number of operations,
it is proportional to the "number of stocks" x "number of rows" in a data set. In this case, it was "12" x "3012" rows = 36,156 operations.

Since Steve  wants to do a little more research for his parents and is gathering data for the entire stock market over the last few years, 
we needed to refactor the code. We assume that there is an unknown number of stocks, the previous code which had a defined array and looped over a number of stocks
is no longer valid. Therefore, we had to construct an algorythms which only looped over the number of rows once, performing necessary tests, incrementing the number of stocks as necessary
and assigning the appropriate values into the array. In terms on efficiency and number of operations,
this new algorythm is proportional only to the number of rows in the data set, in this case of test data set 3012 only.


