Attribute VB_Name = "Module1"
Sub StockData()
    
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    
        WS.Activate
        Lastrow = Cells(Rows.Count, 1).End(xlUp).row

        Cells(1, 9).Value = "ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock vol"
        
        'declare variables
        Dim openP As Double
        Dim closingP As Double
        Dim yrChg As Double
        Dim ticker As String
        Dim PercentChange As Double
        Dim vol As Double
        Dim row As Double
        
        'initialize variables
        row = 2
        vol = 0
        openP = Cells(2, 3).Value
        
        For i = 2 To Lastrow
        
            'check if next ticker equals to current ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         
                ticker = Cells(i, 1).Value
                Cells(row, 9).Value = ticker
                
                closingP = Cells(i, 6).Value
                yrChg = closingP - openP
                Cells(row, 10).Value = yrChg
                
                'percent change
                If (openP = 0 And closingP = 0) Then
                    PercentChange = 0
                ElseIf (openP = 0 And closingP <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = yrChg / openP
                    Cells(row, 11).Value = PercentChange
                    Cells(row, 11).NumberFormat = "0.00%"
                End If
                
                vol = vol + Cells(i, 7).Value
                Cells(row, 12).Value = vol
                row = row + 1
                
                'reset open price and volume
                openP = Cells(i + 1, 3)
                vol = 0
            
            Else
                vol = vol + Cells(i, 7).Value
            End If
        Next i
        
        
        yrChange = Cells(Rows.Count, 9).End(xlUp).row
        
        'conditional formatting
        For j = 2 To yrChange
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
    Next WS
        
End Sub


