# Stock-analysis


## **Overview of Project**: 
In this project, we must analyze the Stock Market Dataset using VBA solution code for our friend Steve. We will be looping through the data to collect stock performances and then refactor the code to make it more efficient by taking few steps, using less memory, and improving our logic of the code to make it easier for future users to read. 

### **Purpose**
The purpose of this project is to refractor the Microsoft VBA code to collect stock information for both years 2017 and 2018, to analyze the stock performances. These stock performances will show which stock is worth investing in. The first analysis in Model 1 of this data was successful, but we have to refractor this code to increase the efficiency of the original code which will be more effective in many ways.

### **Results**
The refactor of stock analysis is done in Module 2. In this module, we have used the code to loop through all the data one time in order to collect the same information we did in module 1. Refactoring this code will help make this VBA script faster. Below are the steps of VBA script and the coding used to perform this analysis. 

**Step 1**

'1a) Create a ticker Index
•	Created the ticketIndex variable and set it equal to zero before iterating all over the rows.

    Dim tickerIndex As Single
    tickerIndex = 0

'1b) Create three output arrays
•	Created the three output arrays: tickerVolumes, tickerStartingPrices and tickerEndingprices. TickerVolume array should be long data type, and tickerStartingPrices should be a single data type.
    
   Dim tickerVolumes(12) As Long
   Dim tickerStartingPrices(12) As Single
   Dim tickerEndingPrices(12) As Single

**Step 2**

''2a) Create a for loop to initialize the tickerVolumes to zero. 

Worksheets(yearValue).Activate
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i
     
 ''2b) Loop over all the rows in the spreadsheet. 
 - Created the code to loop over all the rows. 
    For j = 2 To RowCount
 
 **Step 3**
 
  '3a) Increase volume for current ticker
      
      tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                
         
  '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           
       If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
       tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
             
            End If
            
        'End If
        
  '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
       If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
       '3d Increase the tickerIndex.
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        
        'End If
            End If
        
        
  '3d) 'Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker
        
         If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
   
    
    Next j
    
  **Step 4**
    
  '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
  
        Dim returnValue  As Double
       Worksheets("All Stocks Analysis").Activate
           For i = 0 To 11
           tickerIndex = i
           Cells(4 + i, 1).Value = tickers(tickerIndex)
           Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
           returnValue = (tickerEndingPrices(tickerIndex) - tickerStartingPrices(tickerIndex)) / tickerStartingPrices(tickerIndex)
           Cells(4 + i, 3).Value = returnValue

    Next i

### **Outcomes 2017** 

![VBA_Challenge_2017](https://github.com/Zainak94/stock-analysis/blob/main/VBA_Challenge_2017.PNG)

### **Outcomes 2018**

![VBA_Challenge_2018](https://github.com/Zainak94/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

## **Final Conclusion**

All stocks in 2017 have positive returns except “TERP” with the negative return of 7.20%. Prior to refactoring, the code ran in 1.34375 seconds for the year 2017. After refactoring, the code ran in 0.234375 seconds for the year 2017. This tells us how efficient it makes it for the users to run the code faster than expected. 

All stocks in 2018 have only 2 positive returns of stocks “ENPH” and “RUN”. We can clearly see the performance of the stocks. Prior to refactoring, the code ran in 1.367188 seconds for the year 2018. After refactoring, the code ran in 0.21875 seconds for the year 2018. 

Overall, it gives us clear picture of which stocks will be worth investing in. Refactoring helped us clear and organized data to run it in efficient timely manner. 

## **Summary**


### **Advantages and Disadvantages of Refactoring**

**Pros**

•	Refactoring helps clean the coding process. It will make the code more efficient by only using few steps, it uses less memory and improves the logic of the code which makes     it easier to understand.

•	The logic of code becomes easier to understand when it contains nested conditional and loops. 

•	By adding clear comments and documentation, it makes the data clean and organized.  

**Cons**

•	Refactoring can cause many issues at the same time. Sometimes the data can be too large for not having proper form for existing codes. 

•	If you were to copy and paste the coding, you might run into duplication. 

•	Refactoring can affect the testing of outcomes. There were multiple outcomes received while running the data but they were fixed after careful analysis of the code.

### **Pros & Cons apply to refactoring the original VBA script**

**Pros**

•	Using the original VBA helps in the result of the coding. We have an idea of the result and it helps you understand the coding. All we must do is to update and improve our       code to make it efficient for users.

•	Makes the code cleaner and more well-organized so if you need to make any changes, it will be easier to understand and easy to maintain. 

**Cons**

•	Using the original VBA script can help refractor but at the same time but you might run into some issues. With few errors in coding can change the result. It will be unable     to provide the same outcome as it did previously. 

•	We have to re-check and pay attention to the code refactoring multiple times to run the correct analysis. 

