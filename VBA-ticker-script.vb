Sub ResultForAllTabs()   'Look through all worksheet
                Dim ws As Worksheet
                        For Each ws In ThisWorkbook.Worksheets
                             result ws
                        Next ws
End Sub

'Starting thescript through each indivisual worksheets

Sub result(ws As Worksheet)
        
            'Declaring variables and strings
            Dim tickername As String                         'tickername
            Dim i As Long                                           'Count for looping through all rows
            Dim openvalue As Double                          'Opening value
            Dim closevalue As Double                         'Closing value
            Dim lastrow As Long                               'lastrow count
            Dim j As Long                                         'Counter for displaying
            Dim difference As Double                         'Difference between open and close value
            Dim percentchange As Double                  'Percentage value for difference
            Dim Volume As Double                            'Total volume of stock
            Dim greatVolume As Double                     'for greatest volume calculation
            Dim greattickerVolume As String              'ticker number for greatest volume
            Dim greatIncrease As Double                   'greatest% increase
            Dim greatTickerIncrease As String             'ticker number adjacent to greatest% increase
            Dim greatDecrease As Double                  'greatest% decrease
            Dim greatTickerDecrease As String            'ticker number adjacent to greatest% increase
               
            openvalue = ws.Cells(2, 3).Value               'first opening value
            j = 2
            Volume = 0
            greatVolume = 0
            greatIncrease = 0
            greatDecrease = 0
               
            lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row                'Calculating the last row for looping purpose
               
               ws.Cells(1, 9).Value = ws.Cells(1, 1)                                            'Naming the top rows
               ws.Cells(1, 10).Value = "Yearly Change"
               ws.Cells(1, 11).Value = "Percent change"
               ws.Cells(1, 12).Value = "Total Stock Volume"


'Starting the for loop

    For i = 2 To lastrow
        
                    If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then              'Checking on ticker name conditions
                                closevalue = ws.Cells(i, 6).Value
                                difference = closevalue - openvalue                                         'Calculating the difference
            
                                            If difference < 0 Then                                                                   'Conditions for coloring the cells based on positve and negative values
                                            ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
                                            ws.Cells(j, 11).Interior.Color = RGB(255, 0, 0)
                                                ElseIf difference > 0 Then
                                                    ws.Cells(j, 10).Interior.Color = RGB(81, 255, 13)
                                                    ws.Cells(j, 11).Interior.Color = RGB(81, 255, 13)
                                            End If
                                
                                Volume = Round((Volume + ws.Cells(i, 7)), 0)                        'Calculating the volume and rounding off
                                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                                ws.Cells(j, 10).Value = difference
                                percentchange = (difference / openvalue)                                'Calculating the percentage change
                                ws.Cells(j, 11).Value = FormatPercent(percentchange, 2)
                                ws.Cells(j, 12).Value = Volume
                          
                                            If (greatVolume < Volume) Then                                      'Calculating the greatest ticker volume
                                            greatVolume = Volume
                                            greattickerVolume = ws.Cells(i, 1).Value
                                            End If
                        
                                            If (percentchange > greatIncrease) Then                         'Calculating the greatest ticker percentage increase
                                            greatIncrease = percentchange
                                            greatTickerIncrease = ws.Cells(i, 1).Value
                                            End If
                        
                                            If percentchange < greatDecrease Then                          'Calculating the greatest ticker percentage decrease
                                            greatDecrease = percentchange
                                            greatTickerDecrease = ws.Cells(i, 1).Value
                                            End If
                        
                                    openvalue = ws.Cells(i + 1, 3).Value
                                    j = j + 1
                                    Volume = 0
        
                            Else
                                Volume = Volume + ws.Cells(i, 7)                                                    'Calculating the greatest ticker percentage decrease
                    
                            End If
                
        Next i
        
                            ws.Cells(2, 14) = "Greatest%increase"                                                           'Displaying the greatest % increase
                            ws.Cells(2, 15) = greatTickerIncrease
                            ws.Cells(2, 16) = FormatPercent(greatIncrease, 2)
                            
                            ws.Cells(3, 14) = "Greatest%decrease"                                                          'Displaying the greatest % decrease
                            ws.Cells(3, 15) = greatTickerDecrease
                            ws.Cells(3, 16) = FormatPercent(greatDecrease, 2)
                            
                            ws.Cells(4, 14) = "Greatest total volume"                                                       'Displaying the greatest volume
                            ws.Cells(4, 15) = greattickerVolume
                            ws.Cells(4, 16) = Round(greatVolume, 2)
                            
                            ws.Cells(1, 15) = "Ticker"
                            ws.Cells(1, 16) = "Value"
                            
                             ws.Range("A2:P2").EntireColumn.AutoFit                                                                               'Adjusting the coloum for clear display


        End Sub

