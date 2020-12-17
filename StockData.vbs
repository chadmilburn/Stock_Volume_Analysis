Attribute VB_Name = "Module3"
Sub StockDataFinal()

Dim ws As Worksheet

For Each ws In Worksheets

'set variables--------------------------------------------------------
Dim ticker As String
Dim volume As Double
Dim Summary_Table_Row As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim lrow As Long
    opening_price = ws.Cells(2, 3).Value
    volume = 0
    Summary_Table_Row = 2
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'solve for hard variables
Dim max As Double
Dim min As Double
Dim maxvol As Double
Dim lrowhard As Integer
    
'set column headings--------------------------------------------------
ws.Range("I1").Value = "Ticker"
ws.Range("L1").Value = "Total Volume"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'loop to move through data
For i = 2 To lrow
    
            'while true we are just adding volume under the final else
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                        'when this is false we need to the following
                        ' set and print ticker
                        ticker = ws.Cells(i, 1).Value
                        ws.Range("I" & Summary_Table_Row).Value = ticker
                        'set close price
                        closing_price = ws.Cells(i, 6).Value
                        'figrue and print yearly change
                        yearly_change = closing_price - opening_price
                      
                        ws.Range("J" & Summary_Table_Row).Value = yearly_change
                        ws.Range("J" & Summary_Table_Row).NumberFormat = "###.00"
                        'format color for year change
                        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                        ElseIf ws.Range("J" & Summary_Table_Row).Value = 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.Color = wbwhite
                        Else
                        ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
                        End If
                                'nested to figure % change with 0 values
                                If (opening_price = 0 And closing_price = 0) Then
                                   percent_change = 0
                                ElseIf (opening_price = 0 And closing_price <> 0) Then
                                    percent_change = 1
                                Else
                                    percent_change = yearly_change / opening_price
                                    ws.Range("K" & Summary_Table_Row).Value = percent_change
                                    ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                                End If
                        'add final voulme figure and print
                        volume = volume + ws.Cells(i, 7)
                        ws.Range("L" & Summary_Table_Row).Value = volume
                        'reset open price
                        opening_price = ws.Cells(i + 1, 3).Value
                        'adding rows to the summary table
                        Summary_Table_Row = Summary_Table_Row + 1
                        'reset volume
                        volume = 0
                Else
                    volume = volume + ws.Cells(i, 7).Value
            End If
Next i

'start loop to solve hard answer
'move declartions here to search for max after aummary is built
lrowhard = ws.Cells(Rows.Count, 9).End(xlUp).Row
max = Application.WorksheetFunction.max(ws.Range("K2:K" & lrow))
min = Application.WorksheetFunction.min(ws.Range("K2:K" & lrow))
maxvol = Application.WorksheetFunction.max(ws.Range("L2:L" & lrow))

For J = 2 To lrowhard
                'find and print the max % change
                If ws.Cells(J, 11) = max Then
                        ws.Range("Q2").Value = max
                        ws.Range("Q2").NumberFormat = "0.00%"
                        ws.Range("P2").Value = ws.Cells(J, 9)
                End If
                'find and print the min % change
                If ws.Cells(J, 11) = min Then
                        ws.Range("Q3").Value = min
                        ws.Range("Q3").NumberFormat = "0.00%"
                        ws.Range("P3").Value = ws.Cells(J, 9)
                End If
                'find and pring max volume
                If ws.Cells(J, 12) = maxvol Then
                        ws.Range("Q4").Value = maxvol
                        ws.Range("P4").Value = ws.Cells(J, 9)
                End If

Next J

'format column widths to fit data
ws.Columns("I").AutoFit
ws.Columns("J").AutoFit
ws.Columns("K").AutoFit
ws.Columns("L").AutoFit
ws.Columns("O").AutoFit
ws.Columns("P").AutoFit
ws.Columns("Q").AutoFit

Next ws


End Sub

