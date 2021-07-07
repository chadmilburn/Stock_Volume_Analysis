# VBA-challenge
We were presented stock data initially broken down alphabetically. 
I used this data set to build a script that agreggated that Volume of each ticker utilizing loops and nest if statements
I continued to build on this script addig in Yearly Change and % Change with formatting
Final build involved adding a max and min percent change and max volume for each yearly tab

```
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
```
