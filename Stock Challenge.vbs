Sub StockChallenge()

    ' dim everything
    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim vol As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim summary As Integer
    
    ' loop all sheets
    For Each ws In Worksheets
    
        ' find last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' summary table row starts at 2
        summary = 2
        
        ' set starting vol
        vol = 0
        
        ' loop
        For i = 2 To lastrow
            
            ' if ticker different
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                ' retrieve ticker
                ticker = ws.Cells(i, 1).Value
                
                ' input ticker
                ws.Range("I" & summary).Value = ticker
                
                ' retrieve open and close value
                year_open = ws.Cells(i, 3).Value
                year_close = ws.Cells(i, 6).Value
                
                ' calculate yearly change
                yearly_change = year_close - year_open
                
                ' input yearly change
                ws.Range("J" & summary).Value = yearly_change
                
                    ' confitional formatting for highlight
                    If ws.Range("J" & summary).Value >= 0 Then
                    
                        ' highlight in green
                        ws.Range("J" & summary).Interior.ColorIndex = 4
                        
                    ' if negative
                    Else
                    
                        ' highlight in red
                        ws.Range("J" & summary).Interior.ColorIndex = 3
                        
                    End If
                
                ' calculate percent change
                percent_change = yearly_change / year_open
                
                ' input percent change
                ws.Range("K" & summary).Value = percent_change
                
                ' show percent change in %
                ws.Range("K" & summary).NumberFormat = "0.00%"
                
                ' calculate total vol
                vol = vol + ws.Cells(i, 7).Value
                
                ' input total vol
                ws.Range("L" & summary).Value = vol
                
                ' add a row to summary table
                summary = summary + 1
                
                ' reset vol
                vol = 0
            
            ' if ticker the same
            Else
            
                ' calculate total vol
                vol = vol + ws.Cells(i, 7).Value
            
            End If
        
        Next i
        
        ' set headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' dim greatest
        Dim max_increase As Double
        Dim max_decrease As Double
        Dim max_vol As Double
        
        ' find max
        max_increase = Application.WorksheetFunction.max(ws.Range("K:K"))
        max_decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
        max_vol = Application.WorksheetFunction.max(ws.Range("L:L"))
        
        ' loop
        For j = 2 To lastrow
        
            ' check max increase
            If ws.Cells(j, 11).Value = max_increase Then
            
                ' retrieve related values and input
                ws.Cells(2, 17).Value = ws.Cells(j, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
                
            ' check max decrease
            ElseIf ws.Cells(j, 11).Value = max_decrease Then
            
                ' retrieve related values and input
                ws.Cells(3, 17).Value = ws.Cells(j, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
                
                ' check max vol
            ElseIf ws.Cells(j, 12).Value = max_vol Then
            
                ' retrieve related values and input
                ws.Cells(4, 17).Value = ws.Cells(j, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
  
            End If
            
            ' set data format
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            ws.Range("Q4").NumberFormat = "0.00E+00"
            
        Next j
        
        ' set titles
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' set autofit
        ws.Columns("I:Q").AutoFit
    
    Next ws

End Sub

