# stock-analysis
Analyzing stock data using VBA.

# **PROJECT OVERVIEW and BACKGROUND**

In this project, existing code was refactored to improve its overall efficiency with respect to the following criteria:

 * Using fewer steps
 * Using less memory
 * Improving logic
 * Making the code run faster

Previously, the code (herein referred to as “original” or “original code”) was created to evaluate the performance of a number of stocks over two years (2017 and 2018). The script was run to determine the total daily volume and return on each stock for the specified year. The objective was to determine whether or not the stock would be a good investment.

Only twelve stocks were analyzed, which brought forth the question of whether or not the original code was scalable, and whether it would be as efficient with a larger data set.

To test this, the code was refactored, executed, and its performance was evaluated based on its run time.


# **RESULTS**


## *ORIGINAL RUN vs. REFACTORED RUN*

The original run times were as follows:

![alt text](https://github.com/farwaali08/stock-analysis/blob/6d35b88fe9c31036bfc4a7c05df6124c4ff8b27e/2017_original.png)
> 2017 Run

![alt text](https://github.com/farwaali08/stock-analysis/blob/998e0451279473d25c6aae2416c0fcc34890dfc3/2018_original.png)
> 2018 Run


The refactored code fared well in comparison, with a more rapid run time. The run time decreased from tenths of a second to hundredths of a second—almost tenfold (note: the initial run times were not recorded, and were slightly slower, however the dimensions remained the same):


![alt text](https://github.com/farwaali08/stock-analysis/blob/ce74a25775776ab55f5607f06ae93c8efcc97ec2/VBA_Challenge_2017.png)
> 2017 Run


![alt text](https://github.com/farwaali08/stock-analysis/blob/cf5af9aabf46eaa8ab0d8d66697dde8f8af939df/VBA_Challenge_2018.png)
> 2018 Run

# **ANALYSIS and SUMMARY**

A copy of the the original code can be found within [this]https://github.com/farwaali08/stock-analysis/blob/9e81832873fdbd20887042147662163c80a0ce3d/VBA_CHALLENGE1.xlsm workbook. The original macro was preserved to allow comparison. 



```
For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
```
