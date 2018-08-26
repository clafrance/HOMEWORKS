' Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
' You will also need to display the ticker symbol to coincide with the total volume.

Sub homework2_vb_easy()

    Dim total_volumn As Double
    Dim result_row_count As Integer
    Dim starting_row As Long
    
    For Each ws In Worksheets
    
        total_volumn = 0
        starting_row = 2
        result_row_count = 1
    
        'num_of_rows = Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
        num_of_rows = ws.Cells(Rows.Count, "A").End(xlUp).Row

        num_of_rows_result = ws.Cells(Rows.Count, "I").End(xlUp).Row
        ws.Range("I1", "L" & num_of_rows_result).Clear

        ws.Cells(result_row_count, 9).Value = "Ticker"
        ws.Cells(result_row_count, 10).Value = "Totel Stock Volumn"
    
        For i = 2 To num_of_rows
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                result_row_count = result_row_count + 1
                total_volumn = WorksheetFunction.Sum(ws.Range(ws.Cells(starting_row, 7), ws.Cells(i, 7)))
                ws.Cells(result_row_count, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(result_row_count, 10).Value = total_volumn
                starting_row = i + 1
            End If
        Next i
    Next ws
End Sub