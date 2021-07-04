# Stock_analysis
Module 2 Challenge
## Overview of Project
### Background
We already analysed the data. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

### Purpose
In this challenge, we edited the Module 2 solution code to loop through all the data one time in order to collect the information Then, we determined whether refactoring your code successfully made the VBA script run faster.
### Information
We have these variables in order (columns):

  Ticker - to summarize
  
  Date
  
  Open (price)
  
  High (price)
  
  Low (price)
  
  Close (price) - to analyse
  
  Adj Close (price)
  
  Volume (of transactions) - to analyse

We have 12 tickers of information, each one describes the prices of closure per day during 2017 and 2018 of that stocks.
We need to analyse 3 variables for each ticker presented in this data base: **Volume** (the sum of volume on each day for the ticker), **Starting price** and **Ending price**

Finally we'll present **Total Daily Volume** and **Return** wich is Ending price / Starting price -1

## Results
After doing the analysis of both years, we can see in 2017 these stocks gave us better returns than 2018.
In few words, most of the stocks were down in 2018 except RUN.

![Analysis](https://user-images.githubusercontent.com/85086918/124402017-f543f500-dcf2-11eb-97cb-d3bebc91b6d8.png)

```
Technical Notes:
Ir order to read all the tickers and make calculations, we need to summarize everything by ticker.
All information is in order by ticker and by date, this is important to read the information and make good calculations of Starting and Endind prices.

These are the tickers:

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
    
Then we need nested loops to read all rows of information fo each ticker:
### For each ticker
##### For each row
####### We read the ticker, the volume and the price

Then we make comparations between tickers, last one and this one, to know if this is the first row for the ticker and this one and next row to know if this row is the last for the ticker.

First comparation **to know if the row belongs to this ticker**:
If Cells(i, 1) = tickers(tickerIndex) Then
            
  tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
End If

Second comparation **to know if the row is the first one on this ticker:
If Cells(i, 1).Value <> Cells(i - 1, 1).Value And Cells(i, 1) = tickers(tickerIndex) Then
                
   tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
End If

Third comparation **to know if the row is the last one on this ticker:
If Cells(i, 1).Value <> Cells(i + 1, 1).Value And Cells(i, 1) = tickers(tickerIndex) Then
                
   tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
End If
```

Then, we need to present the information.

For 2017:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/85086918/124402603-f414c700-dcf6-11eb-85d3-14b9419d69d8.png)

For 2018:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/85086918/124402613-01ca4c80-dcf7-11eb-99d1-f7ad2476fefb.png)

**Conclusion: Those stocks are not the best to invest. History tell us the return is not positive in 2018 and we have no data to know if this could change.**

## Summary

**In my opinion the code must be refactorized once you finish the first version aviable.**

For that purpose we need to understad the importance of have comments in each of our logical of coding (algoritm).
If we can read those comments we will understand all the steps in our code and then we can:
- Review if the logical is one of the best ways to solve the objective of the code
- Review each step to organize the best ideas to approach this code
- Review the way we present the data or results
- Review the time our algoritm solve the problem, just in some cases is necessary (so much data or slow computer)

Now, if you're working with partners, they will be able to understand and improve your code if it's necessary.

Maintenance
If this code will work for sometime, we need to review it and refactorize, tech is that way, always improving things.

By Raquel Valdez
