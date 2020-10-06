Attribute VB_Name = "Module1"
Sub Stock_market():
    Dim Tickername As String
    
    Dim Tickertotal As Double
        
    Dim SummaryRow As Long
        
    Dim openprice As Double
        
    Dim closeprice As Double
    Dim Tickervolume As Double
    Dim PercentIncrease As Double
    Dim PercentDecrease As Double
    Dim LargeValue As LongLong
    
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        ' MsgBox LastRow
         
        For Each ws In Worksheets
            ws_name = ws.Name
            openprice = ws.Cells(2, 3).Value
            
            ws.Range("K1").Value = "Ticker"
            ws.Range("L1").Value = "Yearly Change"
            ws.Range("M1").Value = "Percent Change"
            ws.Range("N1").Value = "Total Stock Volume"
            ws.Range("R2").Value = "Greatest % Increase"
            ws.Range("R3").Value = "Greatest % Decrease"
            ws.Range("R4").Value = "Greatest Total Volume"
            ws.Range("S1").Value = "Ticker"
            ws.Range("T1").Value = "Value"
                SummaryRow = 2
            
                For i = 2 To LastRow
                    Tickername = ws.Cells(i, 1).Value
                    NextTicker = ws.Cells(i + 1, 1).Value
                        
                    If NextTicker <> Tickername Then
                               
                        closeprice = ws.Cells(i, 6).Value
                        yearlychange = closeprice - openprice
                            ws.Range("L" & SummaryRow).Value = yearlychange
                                If yearlychange > 0 Then
                                    ws.Range("L" & SummaryRow).Interior.ColorIndex = 4
                                Else
                                    ws.Range("L" & SummaryRow).Interior.ColorIndex = 3
                                End If
                                 
                        
                            If openprice = 0 Then
                                ws.Range("M" & SummaryRow).Value = 0
                            Else
                                PercentChange = yearlychange / openprice
                                ws.Range("M" & SummaryRow).Value = Format(PercentChange, "Percent")
                                
                            End If
                            
                        openprice = ws.Cells(i + 1, 3).Value
                        
                        Tickername = ws.Cells(i, 1).Value
                        ws.Range("K" & SummaryRow).Value = Tickername
                                            
                        Tickertotal = Tickertotal + ws.Cells(i, 7).Value
                            ws.Range("N" & SummaryRow).Value = Tickertotal
                                            
                        SummaryRow = SummaryRow + 1
                   
                        Tickertotal = 0
                        
                    ElseIf ws.Cells(i + 1, 1).Value = Cells(i, 1).Value Then
                        
                        Tickertotal = Tickertotal + ws.Cells(i, 7).Value
                        
                    End If
                    
                Next i
                
                               
                    
            Next ws
    

End Sub

Sub Analysis():
    
  For Each ws In Worksheets
            ws_name = ws.Name
    

      ws.Range("T2") = Application.WorksheetFunction.Max(ws.Range("M"))
      
      
      
    Next ws

                    
                        'PercentIncrease = ws.Range("T2").Value
                    'ws.Application.WorksheetFunction.Min(Range("K:L")) = PercentDecrease
                    'ws.Application.WorksheetFunction.Max(Range("K:L")) = LargeValue
                    


End Sub

