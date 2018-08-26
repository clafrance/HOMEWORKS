' Create a script that will loop through all the stocks and take the following info.

'   -Yearly change from what the stock opened the year at to what the closing price was.
'   -The percent change from the what it opened the year at to what it closed.
'   -The total Volume of the stock
'   -Ticker symbol

' You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub homework2_vb_moderate()

    Dim total_volumn As Double
    Dim result_row_count As Integer
    Dim starting_row As Long
    
    For Each ws In Worksheets

        total_volumn = 0
        starting_row = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Totel Stock Volumn"
    
        'num_of_rows = Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
        num_of_rows = ws.Cells(Rows.Count, "A").End(xlUp).Row

        ws.Range("G1", "G" & num_of_rows).Clear
    
        'MsgBox ("number of rows: " & num_of_rows)
        result_count = 2
    
        For i = 2 To num_of_rows
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                total_volumn = WorksheetFunction.Sum(ws.Range(Cells(starting_row, 7), Cells(i, 7)))
                ' total_volumn = WorksheetFunction.Sum(ws.Range("G" & starting_row, "G" & i))
                'MsgBox ("total: " & total_volumn)
                ws.Cells(result_count, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(result_count, 10).Value = total_volumn
                result_count = result_count + 1
                starting_row = i + 1
            End If
        Next i
    Next ws
End Sub
