Attribute VB_Name = "Module1"
Sub stockcount()

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
  Dim ticker As String
  Dim lastrow As Long
  Dim OpenPrice As Double
  Dim Closeprice As Double
  Dim percentage As Variant
  Dim j As Integer
  Dim YearlyChange As Double
  Dim TotalVolume As Variant
    
  'Initialize the variables
   TotalVolume = 0
   j = 2
   
   
   'Find the lastrow
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
   
   ws.Range("I1").Value = "ticker"
   ws.Range("J1").Value = "Yearly Change"
   ws.Range("K1").Value = "Percentage Change"
   ws.Range("L1").Value = "Total Volume"
   
   OpenPrice = ws.Cells(2, "C")
   'Loop through the column A
   For i = 2 To lastrow
    If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        ticker = ws.Cells(i, "A").Value
        Closeprice = ws.Cells(i, "F").Value
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        
        'calculate yearly change and percentage
        YearlyChange = Closeprice - OpenPrice
        If (OpenPrice <> 0) Then
            percentage = ((Closeprice - OpenPrice) / OpenPrice)
        Else
           percentage = 0
        End If
        'Print on column I , J , K
        ws.Range("I" & j).Value = ticker
        ws.Range("J" & j).Value = YearlyChange
        ws.Range("K" & j).Value = FormatPercent(percentage, 2)
        If (percentage >= 0) Then
            ws.Range("K" & j).Interior.ColorIndex = 4
        Else
            ws.Range("K" & j).Interior.ColorIndex = 3
        End If
        ws.Range("L" & j).Value = TotalVolume
        'inititalize for next ticker
        TotalVolume = 0
        j = j + 1
        OpenPrice = ws.Cells(i + 1, "C").Value
    Else
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    End If
   Next i
   Next ws
End Sub
