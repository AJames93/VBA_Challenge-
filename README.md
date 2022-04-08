# VBA_Challenge-
HW2
Sub Stock_Market_Analysis()

Dim ws As Worksheet

'Walk through each worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

'Label the "Header" of required columns
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Value"

'Dim all necessary inputs
    Dim Ticker As Integer
    Dim Price As Double
    Dim Opening As Long
    Dim Closing As Long
    Dim Annual_Change As Double
    Dim Percent_Change As Double
    
    'Designate lastrow in Data
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        Price = 0
        Ticker = 2
        Opening = 2
    
            For i = 2 To lastrow
            
'Verify the ticker column has generated
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    Cells(Ticker, 9).Value = Cells(i, 1).Value
 
'Populate the Annual_Change column data
    Annual_Change = Cells(i, 6) - Cells(Opening, 3)
    Cells(Ticker, 10).Value = Annual_Change

'Populate the Percent_Change column data
    Percent_Change = Cells(Ticker, 10).Value / Cells(Opening, 3).Value
    Cells(Ticker, 11).NumberFormat = "0.00%"

'Once ticket changes, it will begin to check for the next continuous ticker
    Opening = i + 1
    Cells(Ticker, 11).Value = Percent_Change

'Populates Total Stock Volume in column L2
    Cells(Ticker, 12).Value = Price

    Price = Price + Cells(i, 7).Value
    Ticker = Ticker + 1
    Price = 0

    Else
        Price = Price + Cells(i, 7).Value
    
    
End If
 
    Next i
    
    For i = 2 To lastrow
    If (Cells(i, 10).Value > 0) Then
        Cells(i, 10).Interior.ColorIndex = 4
    Else
        Cells(i, 10).Interior.ColorIndex = 3
        
End If
    Next i
    
                
Next ws

End Sub

