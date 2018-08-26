' Create a script that will loop through all the stocks and take the following info.

'   -Yearly change from what the stock opened the year at to what the closing price was.
'   -The percent change from the what it opened the year at to what it closed.
'   -The total Volume of the stock
'   -Ticker symbol

' You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub homework2_vb_moderate()

    Dim total_volumn As Double
    Dim num_of_rows As Long
    Dim result_row_count As Long
    Dim starting_row As Long
    Dim percent_change As Double
    Dim open_price As Double
    Dim close_price As Double
    
    For Each ws In Worksheets
        
        'num_of_rows = Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
        num_of_rows = ws.Cells(Rows.Count, "A").End(xlUp).Row
        result_row_count = 1
        starting_row = 2

        'num_of_rows_result = ws.Cells(Rows.Count, "I").End(xlUp).Row
        'ws.Range("I1", "N" & num_of_rows).Clear

        ws.Cells(result_row_count, 9).Value = "Ticker"
        ws.Cells(result_row_count, 10).Value = "Yearluy Change"
        ws.Cells(result_row_count, 11).Value = "Percent Change"
        ws.Cells(result_row_count, 12).Value = "Totel Stock Volumn"
        ws.Cells(result_row_count, 13).Value = "Open Price"
        ws.Cells(result_row_count, 14).Value = "Close Price"
    
        For i = 2 To num_of_rows

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

                result_row_count = result_row_count + 1
                total_volumn = WorksheetFunction.Sum(ws.Range(ws.Cells(starting_row, 7), ws.Cells(i, 7)))
                
                open_price = ws.Cells(starting_row, 3).Value
                close_price = ws.Cells(i, 6).Value
                      
                If open_price = 0 Then
                    For j = starting_row To i
                        If ws.Cells(j, 3).Value <> 0 Then
                            open_price = ws.Cells(j, 3).Value
                            Exit For
                        End If
                    Next j
                End If
                
                If close_price = 0 Then
                    For k = i To starting_row Step -1
                        If ws.Cells(k, 6).Value <> 0 Then
                            close_price = ws.Cells(k, 6).Value
                            Exit For
                        End If
                    Next k
                End If

                yearly_change = close_price - open_price
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If

                ws.Cells(result_row_count, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(result_row_count, 10).Value = yearly_change
                If percent_change > 0 Then
                    ws.Cells(result_row_count, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(result_row_count, 11).Interior.ColorIndex = 3
                End If
                ws.Cells(result_row_count, 11).Value = Format(Str(percent_change), "Percent")
                ws.Cells(result_row_count, 12).Value = total_volumn
                ws.Cells(result_row_count, 13).Value = open_price
                ws.Cells(result_row_count, 14).Value = close_price
                           
                starting_row = i + 1
            End If
        Next i
    Next ws
End Sub