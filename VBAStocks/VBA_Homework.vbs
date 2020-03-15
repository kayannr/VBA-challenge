Sub StockData()


Dim ws As Worksheet
Dim r As Double
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim Total_Stock_Vol As Double
Dim summary_row As Double
Dim opn As Double
Dim cls As Double
Dim LastRow As Double
Dim min_percent As Double
Dim max_percent As Double
Dim max_vol As Double
Dim maxticker As String
Dim minticker As String
Dim maxvolticker As String


For Each ws In Worksheets 'run script on every sheet
ws.Range("I" & 1).Value = "Ticker"
ws.Range("J" & 1).Value = "Yearly Change"
ws.Range("K" & 1).Value = "Percent Change"
ws.Range("L" & 1).Value = "Total Stock Volume"

ws.Range("P" & 1).Value = "Ticker"
ws.Range("Q" & 1).Value = "Value"
ws.Range("O" & 2).Value = "Greatest % Increase"
ws.Range("O" & 3).Value = "Greatest % Decrease"
ws.Range("O" & 4).Value = "Greatest Total Volume"

max_percent = 0
min_percent = 0
max_vol = 0
summary_row = 2
Total_Stock_Vol = 0
opn = ws.Cells(2, 3).Value
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For r = 2 To LastRow
        If ws.Cells(r + 1, 1).Value = ws.Cells(r, 1).Value Then
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(r, 7).Value

        ElseIf ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            ticker = ws.Cells(r, 1).Value
            ws.Range("I" & summary_row).Value = ticker
            cls = ws.Cells(r, 6).Value
            ws.Range("J" & summary_row).Value = cls - opn    'yearly change output
            If (ws.Range("J" & summary_row).Value <= 0) Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(summary_row, 10).Interior.ColorIndex = 4
            End If
        
            If opn <> 0 Then
                percent_change = ((cls - opn) / opn) * 100
                'ws.Range("K" & summary_row).NumberFormat = "0.00%" 'convert to perentage
            'ElseIf
                 'MsgBox ("Error. Fix <open> field manually and save the spreadsheet.")
            End If
            
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(r, 7).Value

            ws.Range("K" & summary_row).Value = (CStr(percent_change) & "%")

            ws.Range("L" & summary_row).Value = Total_Stock_Vol

            If (percent_change > max_percent) Then
                    max_percent = percent_change
                    maxticker = ws.Range("I" & summary_row).Value
            ElseIf (percent_change < min_percent) Then
                    min_percent = percent_change
                    minticker = ws.Range("I" & summary_row).Value
            End If
            
            If (Total_Stock_Vol > max_vol) Then
                    max_vol = Total_Stock_Vol
                    maxvolticker = ws.Range("I" & summary_row).Value
            End If

            summary_row = summary_row + 1
            opn = ws.Cells(r + 1, 3).Value
            Total_Stock_Vol = 0

        End If

    Next r

    ws.Range("P2").Value = maxticker
    ws.Range("P3").Value = minticker
    ws.Range("P4").Value = maxvolticker
    ws.Range("Q4").Value = max_vol
    ws.Range("Q2").Value = (CStr(max_percent) & "%")
    ws.Range("Q3").Value = (CStr(min_percent) & "%")
  
Next ws

End Sub