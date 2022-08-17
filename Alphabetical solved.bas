Attribute VB_Name = "Module1"
'Steps
'6 WS being A, B,C ,D,E F
'Create Column header ticker, yearly change,percentage change ,total stock volume
'1. The ticker symbol
'2. Yearly change from opening price at the beginning of a given
'year to the closing price at the end of that year


'-------------------------------------------------------
Sub Stock()
'Loop through All Sheets
For Each ws In Worksheets

Dim WorksheetName As String

'Set the ticker as variable

'Define the variables to be used
Dim ticker As String
Dim volume As Double
Dim price_open As Double
Dim price_close As Double
Dim lastrow As Long

Dim j As Integer

'For Each ws In Worksheet


'Find the Lastrow on each worksheet
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  
'Add the word ticker to the column header J
ws.Cells(1, 10).Value = "Ticker"
'Add the word ticker to the column header K
ws.Cells(1, 11).Value = "Yearly change"
'Add the word ticker to the column header L
ws.Cells(1, 12).Value = "Percentage change"
'Add the word ticker to the column header M
ws.Cells(1, 13).Value = "Total Stock Volume"


j = 2

price_open = ws.Cells(2, 3).Value


'Loop through all ticker
For i = 2 To lastrow

'Get ticker by comparing next row

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

volume = ws.Cells(i, 7).Value + volume

'Set the ticker name
ticker = ws.Cells(i, 1).Value

price_close = Cells(i, 6).Value


'print summary table
ws.Cells(j, 10).Value = ticker
ws.Cells(j, 11).Value = price_close - price_open
ws.Cells(j, 12).Value = (price_close - price_open) / price_open

'print volume
ws.Cells(j, 13).Value = volume

price_open = ws.Cells(i + 1, 3).Value

j = j + 1

volume = 0


Else
volume = ws.Cells(i, 7).Value + volume



End If




Next i
'Add the percentage

For i = 2 To lastrow

For j = 12 To 12
ws.Cells(i, j).NumberFormat = "0.0%"

 
 Next j
 Next i
 


Next ws


End Sub

