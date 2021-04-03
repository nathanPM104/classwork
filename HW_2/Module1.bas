Attribute VB_Name = "Module1"
Sub StockHW()

'Set Variables

Dim Stock_Ticker As String
Dim Total_Volume As Double
Dim Summary_Table_Row As Integer
Dim Yr_Open As Long
Dim Yr_Close As Double
Dim Yr_Delta As Double



'Set initial starting points
Summary_Table_Row = 2
Total_Volume = 0
x = 0
Yr_Open = Cells(2, 3).Value
lastrow = Range("A1").End(xlDown).Row



    'Loop through stock tickers
    For i = 2 To lastrow
    
    
     'Check Cells if they have the same stock symbol
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
     
     'Set Intial Starting points for Stock_Ticker
     Stock_Ticker = Cells(i, 1).Value
    
     Yr_Close = Cells(i, 6).Value
     
     Total_Volume = Total_Volume + Cells(i, 7).Value
     
     Yr_Delta = Yr_Close - Yr_Open
     
     
     'Printing Stick_Ticker
     Cells(Summary_Table_Row, 9).Value = Stock_Ticker
     
     Cells(Summary_Table_Row, 10).Value = Yr_Delta
        If Yr_Delta > 0 Then
        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
        End If
        
     Cells(Summary_Table_Row, 11).Value = Yr_Delta / Yr_Open
     Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
     
     Cells(Summary_Table_Row, 12).Value = Total_Volume
     
     Summary_Table_Row = Summary_Table_Row + 1
     
     Total_Volume = 0
     Yr_Open = Cells(i + 1, 3).Value
       
        
        Else
        Total_Volume = Total_Volume + Cells(i, 7).Value
        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
        

    End If
  Next i


  
  
End Sub
