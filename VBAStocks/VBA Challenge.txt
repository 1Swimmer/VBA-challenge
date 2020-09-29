Sub superstock()

For Each ws In Worksheets

'variables

Dim summaryrow As Integer
Dim total_volume As Double
Dim ticker As String
Dim year_change, percent_change As Double

'Initial value of variables

 total_volume = 0
 summaryrow = 2
 open_price = ws.Cells(2, 3).Value
 
'Summary table and colummn headers

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'Loop to find thru the rows

 For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
 
'Loop code
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
            year_change = ws.Cells(i, 6).Value - open_price
            If open_price = 0 Then
                 percent_change = 0
            Else
                 percent_change = year_change / open_price
            End If
            If year_change < 0 Then
            
'color red
            
            ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
        Else
            
'color Green
            
            ws.Cells(summaryrow, 10).Interior.ColorIndex = 4
        End If
        
'Output of summary table

            ws.Range("I" & summaryrow).Value = ticker
            ws.Range("L" & summaryrow).Value = total_volume
            ws.Range("J" & summaryrow).Value = year_change
            ws.Range("K" & summaryrow).Value = percent_change
            ws.Range("K" & summaryrow).NumberFormat = "0.00%"
            total_volume = 0
            summaryrow = summaryrow + 1
            open_price = ws.Cells(i + 1, 3).Value
        Else
            year_change = ws.Cells(i, 6).Value - open_price
            total_volume = total_volume + Cells(i, 7).Value
       End If
      Next i
    lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    Dim max_percentage As Double
    Dim min_percentage As Double
    Dim max_volume_total As Double
    
'table for maximun and minimun value

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'maximun value of percent change

    max_percentage = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))

'min value of pecent change
    min_percentage = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))

'max value of total Stock volume
    max_volume_total = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))

'Getting value for greatest % Increase
    ws.Range("Q2").Value = max_percentage
    

'% Format
    ws.Range("Q2").NumberFormat = "0.00%"
    
'Getting value for greatest % decrease
    ws.Range("Q3").Value = min_percentage
    
'% Format
    ws.Range("Q3").NumberFormat = "0.00%"
    
'Getting value for greatest total volume
    ws.Range("Q4").Value = max_volume_total
    
'Indetifying & finding ticker maximun percentage, minimun percentage and for maximmun volume value.
    For j = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
    If ws.Cells(j, 11) = max_percentage Then
        ws.Range("p2").Value = ws.Cells(j, 9).Value
    ElseIf ws.Cells(j, 11) = min_percentage Then
        ws.Range("p3").Value = ws.Cells(j, 9).Value
    ElseIf ws.Cells(j, 12) = max_volume_total Then
        ws.Range("p4").Value = ws.Cells(j, 9).Value
    End If
  Next j
 Next ws
End Sub
    



