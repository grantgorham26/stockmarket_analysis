Attribute VB_Name = "Module2"
Sub tickername()
For Each ws In Worksheets

'create variables
Dim ticker_name As String
Dim counter As LongLong
Dim totalvol As LongLong
Dim yearly As Double
Dim percent As Double
Dim lr As Long
Dim openprice As Double
Dim closeprice As Double
Dim great As Double
Dim least As Double
Dim maxvol As LongLong
Dim i As Long
Dim ticker_table As Integer


' create headers and format them
ws.Range("I1").Value = "Ticker"
ws.Range("I1").Font.Bold = True

ws.Range("L1").Value = "Total Volume"
ws.Range("L1").Font.Bold = True

ws.Range("J1").Value = "Yearly Change"
ws.Range("J1").Font.Bold = True

ws.Range("K1").Value = "Percent Change"
ws.Range("K1").Font.Bold = True

ws.Range("N2").Value = "Greatest % Increase:"
ws.Range("N2").Font.Bold = True

ws.Range("N3").Value = "Greatest % Decrease:"
ws.Range("N3").Font.Bold = True

ws.Range("N4").Value = "Greatest Total Volume:"
ws.Range("N4").Font.Bold = True

'assign variables
lr = ws.Cells(Rows.count, 1).End(xlUp).Row




totalvol = 0

ticker_table = 2



'loop through all the tickers
For i = 2 To lr
If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

openprice = ws.Cells(i, 3).Value

'check if we are still in same ticker
ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

closeprice = ws.Cells(i, 6).Value

'set ticker name
 ticker_name = ws.Cells(i, 1).Value
  
 'subtract final day close from first day open
 'need to figure out a variable for 252 is the total amount of days the stock market is open
 ' last day close - first day open
 yearly = closeprice - openprice
 
 'divide yearly change by first day of open price
 'need to figure out how to format as a percentage
 'yearly/firstdayopen
 percent = yearly / openprice
 

 
 'add total vol
 totalvol = totalvol + ws.Cells(i, 7).Value
 
 'paste ticker name into column
 ws.Range("I" & ticker_table).Value = ticker_name
 
 'paste yearly change into column
 ws.Range("J" & ticker_table).Value = yearly
  
 
 
  
 'paste percent change into column
 ws.Range("K" & ticker_table).Value = percent
 ws.Range("K" & ticker_table).NumberFormat = "0.00%"
 'paste total vol in column
ws.Range("L" & ticker_table).Value = totalvol

 'add row to ticker table
 ticker_table = ticker_table + 1
 
 'reset total vol counter
 
 totalvol = 0
 

 Else
 totalvol = totalvol + ws.Cells(i, 7).Value

 End If


 Next i

'calculations for max min
'use match function to match variable to ticker

 great = WorksheetFunction.max(ws.Range("K2:K" & lr).Value)
ws.Range("P2").Value = great
ws.Range("P2").NumberFormat = "0.00%"
maxticker = WorksheetFunction.Match(great, ws.Range("K2:K" & lr), 0)
ws.Range("O2") = ws.Cells(maxticker + 1, 9).Value

least = WorksheetFunction.Min(ws.Range("K2:K" & lr).Value)
ws.Range("P3").Value = least
ws.Range("P3").NumberFormat = "0.00%"
minticker = WorksheetFunction.Match(least, ws.Range("K2:K" & lr), 0)
ws.Range("O3") = ws.Cells(minticker + 1, 9).Value

maxvol = WorksheetFunction.max(ws.Range("L2:L" & lr).Value)
ws.Range("P4").Value = maxvol
maxvolticker = WorksheetFunction.Match(maxvol, ws.Range("L2:L" & lr), 0)
ws.Range("O4") = ws.Cells(maxvolticker + 1, 9).Value

 'formatting loop
Dim j As Long
For j = 2 To lr
If ws.Cells(j, 10).Value > 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(j, 10).Value < 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 3
End If
Next j
Next ws

End Sub





