Attribute VB_Name = "Module1"
Sub tickerStock()
 
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
    ws.Activate
    
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).row

        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quaterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
      
        Dim open_price As Double
        Dim close_price As Double
        Dim quaterly_change As Double
        Dim ticker As String
        Dim percent_change As Double
        
        Dim volume As Double
        Dim row As Double
        Dim column As Integer
        
        volume = 0
        row = 2
        column = 1
       
       
        open_price = ws.Cells(2, column + 2).Value
        
        
        For i = 2 To last_row
        
         
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
            
               
                ticker = ws.Cells(i, column).Value
                ws.Cells(row, column + 8).Value = ticker
                close_price = ws.Cells(i, column + 5).Value
                quaterly_change = close_price - open_price
                ws.Cells(row, column + 9).Value = quaterly_change
                    
                    percent_change = quaterly_change / open_price
                    ws.Cells(row, column + 10).Value = percent_change
                    ws.Cells(row, column + 10).NumberFormat = "0.00%"
 
                volume = volume + ws.Cells(i, column + 6).Value
                ws.Cells(row, column + 11).Value = volume
                
                row = row + 1
                
                open_price = Cells(i + 1, column + 2)
                
                volume = 0
                
            Else
                volume = volume + ws.Cells(i, column + 6).Value
            End If
        Next i
        
        
        quaterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        For j = 2 To quaterly_change_last_row
            If (ws.Cells(j, 10).Value > 0 Or ws.Cells(j, 10).Value = 0) Then
                ws.Cells(j, 10).Interior.ColorIndex = 10
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
    
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
       
        For k = 2 To quaterly_change_last_row
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quaterly_change_last_row)) Then
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quaterly_change_last_row)) Then
                ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quaterly_change_last_row)) Then
                ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
            End If
        Next k
        
        ActiveSheet.Range("I:Q").Font.Bold = True
        ActiveSheet.Range("I:Q").EntireColumn.AutoFit
        Worksheets("Q1").Select
        
    Next ws

End Sub


