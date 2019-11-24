Sub VBA_Stocks():

    'Dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim realCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageCHange As Double
    Dim ws As Worksheet
    
    
    
    For Each ws In Worksheets
        'Set values for each Worksheet
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
 
    
    
    'Row titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    

    
    'Get the row number of the last row with data
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    'MsgBox (rowCount)
     
    For i = 2 To RowCount
    
        'If ticker changes, print the results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Stores results in variable
            total = total + ws.Cells(i, 7).Value
            
            
            
            'Handle zero total volume
            If total = 0 Then
            'results
                ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
            Else
            'Find First non zero value
                If ws.Cells(start, 3) = 0 Then
                    For findValue = start To i
                        If ws.Cells(findValue, 3).Value <> 0 Then
                            start = findValue
                            Exit For
                        End If
                    Next findValue
                End If
            
            'Calculate change
            change = (ws.Cells(i, 6) - ws.Cells(start, 3))
            percentChange = Round((change / ws.Cells(start, 3) * 100), 2)
            
            'start of the next stock ticker
            start = i + 1
            
            'print the results to a separate worksheet
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = Round(change, 2)
            ws.Range("K" & 2 + j).Value = "%" & percentChange
            ws.Range("L" & 2 + j).Value = total
            
            'color palattes positives..green negatives..red
            
            If change > 0 Then
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
            ElseIf change < 0 Then
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            Else
                ws.Range("J" & 2 + j).Interior.ColorIndex = 0
            End If
        
        End If
        
        
        'reset the variable for new stock ticker
        total = 0
        change = 0
        j = j + 1
        days = 0
        dailyChange = 0
    
    'If ticker is still the same, add results
    Else
        total = total + ws.Cells(i, 7).Value
    End If
    
        
    
    Next i

Next ws
    
    
End Sub


