Sub tickervalue()

    For Each ws In Worksheets
        Dim summery_table_row As Integer
        Dim first_opening As Double
        Dim volume As Variant

        summery_table_row = 2
        first_opening = ws.Cells(2, 3).Value
        volume = 0
    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
        For i = 2 To lastRow
            Dim symbol_name As String
            Dim change As Double
            Dim percent_change As Double
            
            volume = volume + ws.Cells(i, 7).Value
           
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                symbol_name = ws.Cells(i, 1).Value
                change = ws.Cells(i, 6).Value - first_opening
                
                If first_opening = 0 Then
                     percent_change = change
                Else
                    percent_change = change / first_opening
                End If
                
                ws.Range("i" & summery_table_row).Value = symbol_name
                ws.Range("j" & summery_table_row).Value = change
                If change < 0 Then
                    ws.Range("j" & summery_table_row).Interior.ColorIndex = 3
                Else
                    ws.Range("j" & summery_table_row).Interior.ColorIndex = 4
                End If
                
                ws.Range("k" & summery_table_row).Value = percent_change
                ws.Range("k" & summery_table_row).NumberFormat = "0.00%"
                ws.Range("L" & summery_table_row).Value = volume
               
                
                
                volume = 0
                summery_table_row = summery_table_row + 1
                first_opening = ws.Cells(i + 1, 3).Value
                    
            End If
          
        Next i
        
        lastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
        
        Dim maximumPercentage As Double
        Dim minimumPercentage As Double
        Dim maximumVolume As Variant
        
        
        maximumPercentage = Application.WorksheetFunction.Max(ws.Range("k2:k" & lastRow))
        For h = 2 To lastRow
            If ws.Cells(h, 11) = maximumPercentage Then
                ws.Cells(2, 17).Value = maximumPercentage
                ws.Cells(2, 17).NumberFormat = "0.00%"
                ws.Cells(2, 16).Value = ws.Cells(h, 9).Value
                Exit For
            End If
        Next h
        
        minimumPercentage = Application.WorksheetFunction.Min(ws.Range("k2:k" & lastRow))
        For h = 2 To lastRow
            If ws.Cells(h, 11) = minimumPercentage Then
                ws.Cells(3, 17).Value = minimumPercentage
                ws.Cells(3, 17).NumberFormat = "0.00%"
                ws.Cells(3, 16).Value = ws.Cells(h, 9).Value
                Exit For
            End If
        Next h
        
        maximumVolume = Application.WorksheetFunction.Max(ws.Range("l2:l" & lastRow))
        For h = 2 To lastRow
            If ws.Cells(h, 12) = maximumVolume Then
                ws.Cells(4, 17).Value = maximumVolume
                ws.Cells(4, 16).Value = ws.Cells(h, 9).Value
                Exit For
            End If
        Next h
       
    
      
     Next ws
End Sub
