Attribute VB_Name = "Module5"
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
   
       'Activate data worksheet
    Worksheets("Module All Stocks Analysis").Activate


   '1) Format the output sheet on All Stocks Analysis worksheet
   yearValue = InputBox("For Module, What year would you like to run the analysis on?")
   
   startTime = Timer
    
    Range("A1").Value = "Module All Stocks (" + yearValue + ")"
    
    'Create a header row
       Cells(3, 1).Value = "Ticker"
       Cells(3, 2).Value = "Total Daily Volume"
       Cells(3, 3).Value = "Return"
   
   'Only allow 2017 and 2018 to input
   If yearValue = 2017 Or yearValue = 2018 Then
   
    'Initialize array of all tickers
      Dim tickers(11) As String
       
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
     
       'Initialize array of all tickers
        'Activate data worksheet
       Worksheets(yearValue).Activate
       'Get the number of rows to loop over
       RowCount = Cells(Rows.count, "A").End(xlUp).Row
    
    
      '1a) Create a ticker Index
       Dim tickerIndex As Integer
       tickerIndex = 0
        
      '1b) Create three output arrays
       Dim tickerVolumes(11)  As Long
       Dim tickerStartingPrices(11) As Single
       Dim tickerEndingPrices(11) As Single
    
      ''2a) Create a for loop to initialize the tickerVolumes to zero.
      For t = 0 To 11
          tickerVolumes(t) = 0
      Next t
        
      ''2b) Loop over all the rows in the spreadsheet.
      For j = 2 To RowCount
            ticker = tickers(tickerIndex)
            'MsgBox ("ticker = " & ticker)
            '3a) Increase volume for current ticker
            If Cells(j, 1).Value = ticker Then
                   tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
            End If
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            End If
            
            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                    tickerIndex = tickerIndex + 1
            End If
      Next j
        
     '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
     Worksheets("Module All Stocks Analysis").Activate
     'Reset tickerIndex to 0
     tickerIndex = 0
     'Loop Inserting the data from the array
     For i = 0 To 11
         'Initialize marginReturn
         Dim marginReturn As Double
         marginReturn = 0
         'ticker
         Cells(4 + tickerIndex, 1).Value = tickers(tickerIndex)
         'ticker values
         Cells(4 + tickerIndex, 2).Value = tickerVolumes(tickerIndex)
         'ticker return
         marginReturn = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
         Cells(4 + tickerIndex, 3).Value = marginReturn
         'Depending on return decide what color
         If marginReturn > 0 Then
                Cells(4 + tickerIndex, 3).Interior.Color = vbGreen
         Else
                Cells(4 + tickerIndex, 3).Interior.Color = vbRed
         End If
         
        'Go to next index
         tickerIndex = tickerIndex + 1
    Next i
    
        'Formatting
        Worksheets("Module All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit
     
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
            
    Else
        MsgBox ("Entered Year Not Found!")
        Cells.Clear
    End If
    
 
End Sub


