## Overview of Project

The project was created to help Steve to see the position of  Green Energy stocks in 2017 and 2018 an if are good. We have used the formulas using tickers. we have developed a code but we received a new code to use only one loop and be faster. 


## Results

In fact, after refactured code, the process runs faster, see bellow the code and the explanation of results.

    
    '1a) Create a ticker Index
   
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    Dim tickervolumes(12) As Long
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
      For i = 0 To 11
      tickervolumes(i) = 0
      Next i
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        

        tickervolumes(tickerIndex) = tickervolumes(tickerIndex) + Cells(i, 8).Value

   
        '3b) Check if the current row is the first row with the selected tickerIndex.
   
      If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
      tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
      End If
            
    
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If
      
    
 
        '3d Increase the tickerIndex.
                       
       If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
       tickerIndex = tickerIndex + 1
       End If
   
       Next i
        
   
    
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickervolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
  
  
  With this code, we have only one Loop running, and we are using Index to calculate all Tickers. The result of the calculation were all the same for both line of codes and years as well. To be easier to test, I created 2 bottoms on excel to use a both codes, and show the times.
  
  Screenshot of 2017 using refactured code:
  
  ![image](https://github.com/lfhoepers/stock-analysis/blob/20ae2535df44f90ea37feebe4c57d231ac1589e5/Resources/VBA_Challenge_2017.PNG)
  
  Screenshot of 2017 using old code and the times between both:
  
  ![image](https://user-images.githubusercontent.com/100812079/157569829-73fec5b7-893e-41bc-bc85-1c7838237449.png)
 
  
  Screenshot of 2018 using refactured code:
  
  ![image](https://github.com/lfhoepers/stock-analysis/blob/20ae2535df44f90ea37feebe4c57d231ac1589e5/Resources/VBA_Challenge_2018.PNG)
  
  Screenshot of 2018 using old code and the times between both:
  
  ![image](https://user-images.githubusercontent.com/100812079/157570934-ef4cd802-25a5-447e-8537-b0cc5ac22c69.png)

My final analysis about the results, are 2017 for sure was a better year for this company than 2018. When we have a negative % for almost 100% of tickers.

# Summary

As a Developer, I always try to avoid too much diferent loops. So using index in this case is faster and could be safer to avoid issues. Also the data can increase and refactored code will support better.

**Advantages and Disadvantages of Refactoring**

**Advantages**

- Review the code is always a way to find issues;
- We can find faster ways for the code;
- We can comment better all the code;

**Disadvantages**

- If we don't take care, we can bring different results;
- Could spend too much time;

**Advantages and Disadvantages of the original and Refactored Code**

**Advantages**

Smaller code is: 

- better to maintenance;
- better to understanting of new developers;
- less chance to errors;

**Disadvantages**

- The only disadvantage that I found was the time spent to rewrite the code and test.

I appreciate the opportunity to present this project, I am available for any clarification.

**Luiz Fernando Hoepers**  
###### UofT SCS Data Analytics Boot Camp
