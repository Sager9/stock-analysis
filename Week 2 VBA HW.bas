Attribute VB_Name = "Module1"
Sub VBA():
    
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
    
    
        Dim worksheetname As String
        
        Dim ticker As String
        Dim vol As LongLong
        Dim v As Integer
        v = 2
        Dim start As Double
        Dim closing As Double
        Dim change As Double
        Dim change_p As Double
        
        
        
        Dim lastrow As LongLong
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
            For y = 1 To lastrow
                    'Assine top row
                    If y = 1 Then
                        ws.Cells(1, 10) = "Yearly Change"
                        ws.Cells(1, 9) = "Ticker"
                        ws.Cells(1, 11) = "Percent Change"
                        ws.Cells(1, 12) = "Total Stock Volume"
                        
                        Else
            
                            'assige ticker value
                            ticker = ws.Cells(y, 1)
                            'assine start value
                            If Cells(y, 1) <> Cells(y - 1, 1) Then
                                start = ws.Cells(y, 3)
                            End If
                            
                            
                            If ws.Cells(y, 1).value = ws.Cells(y + 1, 1).value Then
                                vol = vol + ws.Cells(y, 7).value
                            Else
                            'finding change
                            closing = ws.Cells(y, 6)
                            change = closing - start
                            If change < 0 Then
                                ws.Cells(v, 10) = change
                                ws.Cells(v, 10).Interior.Color = RGB(255, 0, 0)
                            Else
                                ws.Cells(v, 10) = change
                                ws.Cells(v, 10).Interior.Color = RGB(0, 255, 0)
                            End If
                            
                            change_p = Round(change / start, 4)
                            
                            If change_p < 0 Then
                                ws.Cells(v, 11) = FormatPercent(change_p)
                                ws.Cells(v, 11).Interior.Color = RGB(255, 0, 0)
                            Else
                                ws.Cells(v, 11) = FormatPercent(change_p)
                                ws.Cells(v, 11).Interior.Color = RGB(0, 255, 0)
                            End If
                            
                            
                            
                            
                            'totals vol and assine ticker
                            vol = vol + ws.Cells(y, 7).value
                            ws.Cells(v, 12) = vol
                            ws.Cells(v, 9) = ticker
                            vol = 0
                            
                            'output on next row
                            v = v + 1
                            End If
                        End If
                                   
                Next y
                
                'Hilights value list
                
                Dim i As Double
                Dim it As String
                Dim d As Double
                Dim dt As String
                Dim value As LongLong
                Dim vt As String
                
                i = 0
                d = 0
                value = 0
                
                lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
                For y = 2 To lastrow
                    'greates increase
                    If ws.Cells(y, 11) > i Then
                        i = ws.Cells(y, 11)
                        it = ws.Cells(y, 9)
                    End If
                    
                    'Greatest decrease
                    If ws.Cells(y, 11) < d Then
                        d = ws.Cells(y, 11)
                        dt = ws.Cells(y, 9)
                    End If
                    
                    'greatest volume
                    If ws.Cells(y, 12) > value Then
                        value = ws.Cells(y, 12)
                        vt = ws.Cells(y, 9)
                    End If
                Next y
                ws.Cells(1, 15) = "Ticker"
                ws.Cells(1, 16) = "Value"
                ws.Cells(2, 14) = "Greatest % Increase"
                ws.Cells(3, 14) = "Greatest % Decrease"
                ws.Cells(4, 14) = "Greatest Total Volume"
                ws.Cells(2, 15) = it
                ws.Cells(2, 16) = FormatPercent(i)
                ws.Cells(3, 15) = dt
                ws.Cells(3, 16) = FormatPercent(d)
                ws.Cells(4, 15) = vt
                ws.Cells(4, 16) = value
            
                    
    Next ws
    
End Sub






