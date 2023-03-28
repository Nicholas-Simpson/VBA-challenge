# VBA-challenge

Here is my code for the VBA challenge 2 homework

Sub stocksData()
    Dim lastrow, yearChange, firstPrice, lastPrice, totalValue, rowCounter As Double
    Dim sheetName As String
    Dim maxPercent, minPercent, totalVolume As Double
    Dim FPCounter As Integer
    
    ## this section is the for loop for looping through each worksheet to create the tables from the homework
    as you can see I set the last row using the formula for auto tracking the last row. the other variables are mostly set to 0 for later use 
    within the code. Setting the rowCounter to 2 allows for the table to start in the second row of the excel sheet. To keep track of the first 
    iteration of each loop through the for loop for i I had to initiate the FPCounter to -1 to always bring the math back to the first iteration 
    of the new ticker calculation. 
    
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        sheetName = ws.Name
        totalValue = 0
        maxPercent = 0
        minPercent = 0
        totalVolume = 0
        rowCounter = 2
        FPCounter = -1
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        For i = 2 To lastrow
            
            ## this section just keeps the running total of the value of the stock we are attempting to sum along with counting which row is first in the 
            set.
            totalValue = ws.Cells(i, 7) + totalValue
            FPCounter = FPCounter + 1
            
            ## here the code gets a little complicated. The statement compares the name of each line to see if we have hit a new stock. It will pass this section 
            if the comparison is showing the line we are on is the same as the line we just left. If the line is different it will enter this section of code.
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                ##this section is just setting the values of the table we are trying to create by using the rowCounter as its number on the new table along with 
                writing the totalValue in the total stock volume section. It also is where we start doing the calculations for the yearly change and preping 
                variables to do percent change in stock.
                ws.Cells(rowCounter, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(rowCounter, 12).Value = totalValue
                firstPrice = ws.Cells(i - FPCounter, 3).Value
                lastPrice = ws.Cells(i, 6).Value
                yearChange = lastPrice - firstPrice
                ws.Cells(rowCounter, 10).Value = yearChange
                
                ## here we are setting the color index of the yearly change section based on its value. If its greater than 0 we make it green if its not we make
                it red. 
                If ws.Cells(rowCounter, 10).Value > 0 Then
                    ws.Cells(rowCounter, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowCounter, 10).Interior.ColorIndex = 3
                End If
                
                ## In this section we are using the aboved mentioned variables to calculate the percent change of the stock and format its print out to a percent.
                ws.Cells(rowCounter, 11).Value = (yearChange / firstPrice)
                ws.Cells(rowCounter, 11).NumberFormat = "0.00%"
                
                ## here is where I think I could deffinatly make some improvements to the code overall. This section currently is just constantly printing out what 
                stock has the highest, lowest and most stock volume. Rather than printing it out for each run through this section of code I realized later that I may
                be more efficient code wise to store these values in variables of use an array to store the values and print them after the first table is created to 
                save on processing power.
                If ws.Cells(rowCounter, 11).Value > maxPercent Then
                    maxPercent = ws.Cells(rowCounter, 11).Value
                    ws.Range("P2").Value = ws.Cells(rowCounter, 9).Value
                    ws.Range("Q2").Value = maxPercent
                    ws.Range("Q2").NumberFormat = "0.00%"
                ElseIf ws.Cells(rowCounter, 11).Value < minPercent Then
                    minPercent = ws.Cells(rowCounter, 11).Value
                    ws.Range("P3").Value = ws.Cells(rowCounter, 9).Value
                    ws.Range("Q3").Value = minPercent
                    ws.Range("Q3").NumberFormat = "0.00%"
                End If
                
                If totalValue > totalVolume Then
                    totalVolume = totalValue
                    ws.Range("P4").Value = ws.Cells(rowCounter, 9).Value
                    ws.Range("Q4").Value = totalVolume
                End If
                
                ## this section just resets our totalValue and FPCounter for the next stock while moving our rowCounter one space as to not overwrite the previous new 
                entery on our table.
                totalValue = 0
                FPCounter = -1
                rowCounter = rowCounter + 1
                
            End If
        
        Next i
        
    Next ws


End Sub
