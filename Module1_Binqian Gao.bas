Attribute VB_Name = "Module1"
Sub test()


Dim last_price, first_price As Double
Dim ws As Worksheet
For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 tickernum = 1
 For rownum = 2 To lastrow
     
    If ws.Cells(rownum, 1).Value = ws.Cells(rownum + 1, 1).Value Then
     cnt = cnt + 1
     sm = sm + ws.Cells(rownum, 7).Value
    ElseIf ws.Cells(rownum, 1).Value <> ws.Cells(rownum + 1, 1).Value Then
     tickernum = tickernum + 1
     last_price = ws.Cells(rownum, 6).Value
     first_price = ws.Cells(rownum - cnt, 3).Value
     ws.Cells(tickernum, 9).Value = ws.Cells(rownum, 1).Value
     ws.Cells(tickernum, 10).Value = last_price - first_price
       If ws.Cells(tickernum, 10).Value <> 0 And first_price <> 0 Then
           ws.Cells(tickernum, 11).Value = (last_price - first_price) / first_price
           ws.Cells(tickernum, 12).Value = sm + ws.Cells(rownum, 7).Value
       ElseIf Cells(tickernum, 10).Value = 0 Or first_price = 0 Then
           ws.Cells(tickernum, 11).Value = 0
           ws.Cells(tickernum, 12).Value = sm + ws.Cells(rownum, 7).Value
       End If
     sm = 0
     cnt = 0
   End If
  Next rownum
  
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  ws.Cells(1, 15).Value = "Ticker"
  ws.Cells(1, 16).Value = "Value"
  ws.Cells(2, 14).Value = "Greatest % Increase"
  ws.Cells(3, 14).Value = "Greatest % Decrease"
  ws.Cells(4, 14).Value = "Greatest Total Volume"
  
lastrow_compare = ws.Cells(Rows.Count, 9).End(xlUp).Row
  numMax = -1
  numMin = 1
  Sumax = 0
  
  For i = 2 To lastrow_compare
   If numMax < ws.Cells(i, 11).Value Then
    numMax = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = numMax
    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
   End If
  Next i
 
  For m = 2 To lastrow_compare
   If numMin > ws.Cells(m, 11).Value Then
    numMin = ws.Cells(m, 11).Value
    ws.Cells(3, 16).Value = numMin
    ws.Cells(3, 15).Value = ws.Cells(m, 9).Value
   End If
    Next m
   
 For n = 2 To lastrow_compare
  If Sumax < ws.Cells(n, 12).Value Then
    Sumax = ws.Cells(n, 12).Value
    ws.Cells(4, 16).Value = Sumax
    ws.Cells(4, 15).Value = ws.Cells(n, 9).Value
   End If
   
 Next n
  
  
Next ws

End Sub

