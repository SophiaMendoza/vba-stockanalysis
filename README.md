# vba-stockanalysis

Sub stock_analysis()
    ' Set dimensions
    Dim total As Double
    Dim rowindex As Long
    Dim change As Double
    Dim columnindex As Integer
    Dim Start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Single
    Dim averageChange As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        columnindex = 0
        total = 0
        change = 0
        Start = 2
        dailyChange = 0
        
        ' Set title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("Q2").Value = "Greatest % Increases"
        ws.Range("Q3").Value = "Greatest % Decreases"
        ws.Range("Q4").Value = "Greatest Total Volume"
        
        'get the row number of the last row with data
        rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        For rowindex = 2 To rowCount
            'if ticker changes then print results
            If ws.Cells(rowindex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
                ' Store results in variables
                total = total + ws.Cells(rowindex, 7).Value
                
                If total = 0 Then
                    ' Print the results
                    ws.Range("I" & 2 + columnindex).Value = ws.Cells(rowindex, 1).Value
                    ws.Range("J" & 2 + columnindex).Value = 0
                    ws.Range("K" & 2 + columnindex).Value = "%" & 0
                    ws.Range("L" & 2 + columnindex).Value = 0
                Else
                    If ws.Cells(Start, 3) = 0 Then
                        For find_value = Start To rowindex
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                Start = find_value
                                Exit For
                            End If
                        Next find_value
                    End If
                    change = (ws.Cells(rowindex, 6) - ws.Cells(Start, 3))
                    percentChange = change / ws.Cells(Start, 3)
                    
                    Start = rowindex + 1
                    
                    ws.Range("I" & 2 + columnindex).Value = ws.Cells(rowindex, 1).Value
                    ws.Range("J" & 2 + columnindex).Value = change
                    ws.Range("J" & 2 + columnindex).NumberFormat = "0.00"
                    ws.Range("K" & 2 + columnindex).Value = percentChange
                    ws.Range("K" & 2 + columnindex).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + columnindex).Value = total
                    
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + columnindex).Interior.ColorIndex = 0
                    End Select
                End If
                
                total = 0
                change = 0
                columnindex = columnindex + 1
                days = 0
                dailyChange = 0
            Else
                ' If ticker is still the same, add results to the total
                total = total + ws.Cells(rowindex, 7).Value
            End If
        Next rowindex
        
        ' Calculate the max and min and place them in a separate part of the worksheet
        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        ' Calculate the row numbers for greatest increases, decreases, and total volume
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("k2:k" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("k2:k" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        
        ' Set the values for the greatest increases, decreases, and total volume
        ws.Range("P2").Value = ws.Cells(increase_number + 1, 1).Value
        ws.Range("P3").Value = ws.Cells(decrease_number + 1, 1).Value
        ws.Range("P4").Value = ws.Cells(volume_number + 1, 1).Value
    Next ws
End Sub
