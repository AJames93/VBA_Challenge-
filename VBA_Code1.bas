Attribute VB_Name = "Module1"
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
    Dim Opening As Double
    Dim Closing As Double
    Dim Annual_Change As Double
    Dim Percent_Change As String
    
'Designate lastrow in Data
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Price = 0
    Ticker = 2
    Opening = Cells(2, 3).Value
    
        For i = 2 To lastrow
            
'Verify the ticker column has generated
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    Cells(Ticker, 9).Value = Cells(i, 1).Value
 
 Closing = Cells(i, 6).Value
 
'Populate the Annual_Change column data
    Annual_Change = Closing - Opening
    Cells(Ticker, 10).Value = Annual_Change

'Populate the Percent_Change column data
     If Opening = 0 Then
            Range("K" & Ticker).Value = Annual_Change / 0.01
        Else
            Range("K" & Ticker).Value = Annual_Change / Opening
        End If
    
'Once ticker changes, it will begin to check for the next continuous ticker
    

'Populates Total Stock Volume in column L2
    Price = Price + Cells(i, 7).Value
    Cells(Ticker, 12).Value = Price
    
    Ticker = Ticker + 1
    Price = 0
    Opening = Cells(i + 1, 3).Value

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

Columns("K").NumberFormat = "0.00%"
    
                
Next ws

End Sub
