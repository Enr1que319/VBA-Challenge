Sub Home()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim te, numrow, numticker, h, l As Integer
Dim ticker, ticker2, name As String
'Dim clo, op, voltot, vol As Long


For j = 1 to 3

    ActiveWorkbook.Worksheets(j).Select

    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"

    Range("A2").Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlYes
    name = ActiveWorkbook.Worksheets(j).name
    numrow = Worksheets(name).UsedRange.Rows.Count

    voltot = 0
    h = 2
    l = 0
    opop = 0

    For k = 2 To numrow

        If opop = 0 Then
            op = Cells(k, 3).Value
        End If
        
        clo = Cells(k, 6).Value
        ticker = Cells(k, 1).Value
        ticker2 = Cells(k + 1, 1).Value
        vol = Cells(k, 7).Value
        
        If ticker = ticker2 And Not IsNull(ticker2) Then
            voltot = voltot + vol
            opop = opop + 1
        Else

            voltot = voltot + vol
            Cells(h, 9) = ticker
            Cells(h, 10) = clo - op

            If op <> 0 Then
                Cells(h, 11) = (clo - op) / op
            Else
                Cells(h, 11) = 0
            End If

            Cells(h, 11).NumberFormat = "0.00%"
            Cells(h, 12) = voltot
        
            If Cells(h, 10) <= 0 Then
                Cells(h, 10).Interior.ColorIndex = 3
            Else
                Cells(h, 10).Interior.ColorIndex = 4 
            End If

            h = h + 1
            opop = 0
            voltot = 0

        End If
    Next k

    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"

    Range("Q2") = Application.WorksheetFunction.Max(Range("K:K"))
    Range("Q4") = Application.WorksheetFunction.Max(Range("L:L"))
    Range("Q3") = Application.WorksheetFunction.Min(Range("K:K"))

    Range("I2").Activate
    count_rows = Range(Selection, Selection.End(xlDown)).Count

    y = 11

    For f = 2 To 4
        For d = 2 To count_rows

            If f = 4 Then
                y = 12
            End If

            If Range("Q" & f).Value = Cells(d, y) Then
                Range("P" & f) = Cells(d, 9).Value
            End If

        Next d
    Next f

    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("I:Q").EntireColumn.AutoFit

Next j

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


