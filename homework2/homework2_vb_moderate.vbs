' Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
' You will also need to display the ticker symbol to coincide with the total volume.

Sub homework2_vb_easy()

    Dim total_volumn As Double
    Dim result_row_count As Integer
    Dim starting_row As Long
    
    For Each ws In Worksheets
        total_volumn = 0
        starting_row = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Totel Stock Volumn"
    
        'num_of_rows = Worksheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
        num_of_rows = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        'MsgBox ("number of rows: " & num_of_rows)
        result_count = 2
    
        For i = 2 To num_of_rows
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                total_volumn = WorksheetFunction.Sum(ws.Range("G" & starting_row, "G" & i))
                yearly_change = ws.Cells(starting_row & "C").value - ws.Range("C" & starting_row, "G" & i).value 

                'MsgBox ("total: " & total_volumn)
                ws.Cells(result_count, 9).Value = ws.Cells(i, 1).Value


                ws.Cells(result_count, 12).Value = total_volumn
                result_count = result_count + 1
                starting_row = i + 1
            End If
        Next i
    Next ws
End Sub
