{\rtf1\ansi\ansicpg1252\cocoartf2706
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub StockPrice()\
\
For Each ws In Worksheets\
\
\
    ws.Range("I:O").Clear\
    ws.Range("I1:L1").Value = Array("Ticker", "Price Change", "Percentage Change", "Volume")\
    ws.Range("N2").Value = "Greatest % Increase"\
    ws.Range("N3").Value = "Greatest % decrease"\
    ws.Range("N4").Value = "Greatest Total Volume"\
    ws.Range("O1:P1").Value = Array("Ticker", "Value")\
    ws.Range("L:L").NumberFormat = "General"\
    \
  'STEP 1 - CALCULATE INDIVIDUAL STOCK TICKERS\
    \
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row\
    \
    For i = 2 To lastrow\
    \
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then\
            ticker = Cells(i, 1).Value\
            opening_price = Cells(i, 3).Value\
            yearly_volume = 0\
            \
        End If\
        \
        yearly_volume = yearly_volume + ws.Cells(i, 7)\
        \
        If ws.Cells(i, 1).Value <> ws.Cells((i + 1), 1) Then\
                closing_price = ws.Cells(i, 6).Value\
                lastrowticker = ws.Cells(Rows.Count, 9).End(xlUp).Row\
                ws.Cells((lastrowticker + 1), 9).Value = ticker\
                ws.Cells((lastrowticker + 1), 10).Value = (closing_price - opening_price)\
                If opening_price <> 0 Then\
                    ws.Cells((lastrowticker + 1), 11).Value = Format(((closing_price - opening_price) / opening_price), "0.00%")\
                End If\
                If ws.Cells((lastrowticker + 1), 10).Value >= 0 Then\
                    ws.Cells((lastrowticker + 1), 10).Interior.Color = vbGreen\
                Else\
                    ws.Cells((lastrowticker + 1), 10).Interior.Color = vbRed\
                End If\
                ws.Cells((lastrowticker + 1), 12).Value = yearly_volume\
\
                \
        End If\
    \
    Next i\
\
Next ws\
\
End Sub\
\
Sub SummaryTicker()\
\
   ' STEP 2 - CALCULATE SUMMARIES\
 For Each ws In Worksheets\
 \
    lastrowticker = ws.Cells(Rows.Count, 9).End(xlUp).Row\
    \
    greatest_increase = 0\
    greatest_increase_ticker = ""\
    greatest_decrease = 0\
    greatest_decrease_ticker = ""\
    greatest_volume = 0\
    greatest_volume_ticker = ""\
    \
    For i = 2 To lastrowticker\
        If ws.Cells(i, 11).Value > greatest_increase Then\
            greatest_increase = ws.Cells(i, 11).Value\
            greatest_increase_ticker = ws.Cells(i, 9).Value\
        End If\
        If ws.Cells(i, 11).Value <= greatest_decrease Then\
            greatest_decrease = ws.Cells(i, 11).Value\
            greatest_decrease_ticker = ws.Cells(i, 9).Value\
        End If\
        If ws.Cells(i, 12).Value > greatest_volume Then\
            greatest_volume = ws.Cells(i, 12).Value\
            greatest_volume_ticker = ws.Cells(i, 9).Value\
        End If\
    \
    Next i\
    \
    ws.Range("O2").Value = greatest_increase_ticker\
    ws.Range("P2").Value = Format(greatest_increase, "0,00%")\
    ws.Range("O3").Value = greatest_decrease_ticker\
    ws.Range("P3").Value = Format(greatest_decrease, "0,00%")\
    ws.Range("O4").Value = greatest_volume_ticker\
    ws.Range("P4").Value = greatest_volume\
 \
 Next ws\
\
End Sub\
\
Sub SummaryAlternative()\
\
' ALTERNATIVE WAY TO FIND THE GREATEST DECREASE, INCREASE AND VOLUME\
\
For Each ws In Worksheets\
\
    lastrowticker = ws.Cells(Rows.Count, 9).End(xlUp).Row\
\
'find greatest decrease\
\
    Application.Volatile\
        With Application.WorksheetFunction\
            DMin = .Min(ws.Range("K2:K" & lastrowticker))\
            lIndex = .Match(DMin, ws.Range("K2:K" & lastrowticker), 0)\
        End With\
        GetAddr = ws.Range("K2:K" & lastrowticker).Cells(lIndex).Address\
        \
    ws.Range("P3").Value = DMin\
    ws.Range("O3").Value = ws.Range("I" & ws.Range(GetAddr).Row).Value\
\
'find greatest increase\
\
    Application.Volatile\
        With Application.WorksheetFunction\
            DMax = .Max(ws.Range("K2:K" & lastrowticker))\
            lIndex = .Match(DMax, ws.Range("K2:K" & lastrowticker), 0)\
        End With\
        GetAddr = ws.Range("K2:K" & lastrowticker).Cells(lIndex).Address\
        \
    ws.Range("P2").Value = DMax\
    ws.Range("O2").Value = ws.Range("I" & ws.Range(GetAddr).Row).Value\
\
'find greatest volume\
\
    Application.Volatile\
        With Application.WorksheetFunction\
            DMaxVol = .Max(ws.Range("l2:l" & lastrowticker))\
            lIndex = .Match(DMaxVol, ws.Range("l2:l" & lastrowticker), 0)\
        End With\
        GetAddr = ws.Range("l2:l" & lastrowticker).Cells(lIndex).Address\
        \
    ws.Range("P4").Value = DMaxVol\
    ws.Range("O4").Value = ws.Range("I" & ws.Range(GetAddr).Row).Value\
\
Next ws\
\
End Sub\
\
}