Attribute VB_Name = "Module1"

Option Explicit
Sub StockSymbol3():

  Dim ws As Worksheet
  For Each ws In Worksheets
  
  Const FIRST_DATA_ROW As Integer = 2
  

  Dim StockSymbol As String
  Dim StockVolume As Double
  Dim lastrow As Long
  Dim lastrow2 As Long
  Dim inputRow As Long
  Dim PercentChange As Double
  Dim Max_Total As Long
  
  Dim Max_Row As Integer
  Dim Min_Row As Integer
  Dim Max_Total_Row As Integer
  Dim MaxPercentage As Double
  Dim MinPercentage As Double
  


  
  
  StockVolume = 0
 
  'Put in Headings
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
   
   ws.Cells(1, 16).Value = "Ticker"
   ws.Cells(1, 17).Value = "Value"
   ws.Cells(2, 15).Value = "Greatest % Increase"
   ws.Cells(3, 15).Value = "Greatest % Decrease"
   ws.Cells(4, 15).Value = "Greatest Total Volume"
   
   'Go to last cell in a column
   
  

  ' Keep track of the location for each stock symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = FIRST_DATA_ROW

  
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
'*******************************************************************************************************************
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    
    Summary_Table_Row = FIRST_DATA_ROW
    OpenPrice = ws.Cells(2, 3).Value
    
For inputRow = FIRST_DATA_ROW To lastrow
    StockVolume = StockVolume + ws.Cells(inputRow, 7).Value
    'Check if we are still in same stock symbol, if it is not
    If (ws.Cells(inputRow + 1, 1).Value <> ws.Cells(inputRow, 1).Value) Then
        StockSymbol = ws.Cells(inputRow, 1).Value
        ClosePrice = ws.Cells(inputRow, 6).Value
        
        'Calculations
        YearlyChange = ClosePrice - OpenPrice
        PercentChange = YearlyChange / OpenPrice
        
        'Output
        ws.Range("I" & Summary_Table_Row).Value = StockSymbol
        ws.Range("j" & Summary_Table_Row).Value = YearlyChange

        ws.Range("k" & Summary_Table_Row).Value = PercentChange
            If PercentChange >= 0 Then
                 ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            Else: ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
        ws.Range("L" & Summary_Table_Row).Value = StockVolume
        
        'Set up for next stock
        StockVolume = 0
        OpenPrice = ws.Cells(inputRow + 1, 3)
        Summary_Table_Row = Summary_Table_Row + 1
    End If
Next inputRow
    
lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
Summary_Table_Row = FIRST_DATA_ROW


MaxPercentage = WorksheetFunction.Max(ws.Range("K2:K" & lastrow2))
MinPercentage = WorksheetFunction.Min(ws.Range("K2:K" & lastrow2))
ws.Range("Q2") = MaxPercentage
ws.Range("Q3") = MinPercentage
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow2))
Max_Row = WorksheetFunction.Match(MaxPercentage, ws.Range("K2:K" & lastrow2), 0)
Min_Row = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & lastrow2), 0)
Max_Total_Row = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & lastrow2), 0)
ws.Range("P2") = ws.Cells(Max_Row + 1, 9)
ws.Range("P3") = ws.Cells(Min_Row + 1, 9)
ws.Range("P4") = ws.Cells(Max_Total_Row + 1, 9)
   
Dim Summary_Table_Row2 As Integer

For Summary_Table_Row2 = FIRST_DATA_ROW To lastrow2
    ws.Range("k" & Summary_Table_Row2).NumberFormat = "0.00%"
Next Summary_Table_Row2
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"

ws.Columns(10).AutoFit
ws.Columns(11).AutoFit
ws.Columns(12).AutoFit
ws.Columns(15).AutoFit
ws.Columns(17).AutoFit

ws.Activate
Next ws
    
    
End Sub


