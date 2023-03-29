Attribute VB_Name = "Formatting"
Sub tableformatting(ws As Worksheet)

'MAIN SCRIPT SETUP
'------------------

Dim lastrow As Long
Dim sumtablelastrow As Long

lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
sumtablelastrow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row


Dim summarytable_header As Range
Set summarytable_header = Range("L1:Z1")

'Formatting the Percentages

Dim formatting_colnum_pct As New Collection
formatting_colnum_pct.Add 14
formatting_colnum_pct.Add 18

Dim i As Variant
For Each i In formatting_colnum_pct
    ws.Columns(i).NumberFormat = "0.00%"
    Next i

'Formatting the Values with color

For i = 2 To sumtablelastrow
    If ws.Cells(i, 13).Value < 0 Then
            ws.Cells(i, 13).Interior.ColorIndex = 3
        Else
            ws.Cells(i, 13).Interior.ColorIndex = 4
        End If
Next i

'formatting the avg dimension checker
Dim yrly_pct_col As Range
Dim avg_pct_col As Range

Set yrly_pct_col = ws.Range("N2:N" & sumtablelastrow)
Set avg_pct_col = ws.Range("U2:U" & sumtablelastrow)

For i = 2 To sumtablelastrow
    If Abs(avg_pct_col.Cells(i, 1).Value) - Abs(yrly_pct_col.Cells(i, 1).Value) <= 0.05 Then
        ws.Cells(i, 22).Value = "MATCH"
    End If
Next i


'including header text with formatting
summarytable_header.Columns(1).Value = "Stock Tickers"
summarytable_header.Cells(2).Value = "Yearly Change"
summarytable_header.Cells(3).Value = "Percent Change"
summarytable_header.Cells(4).Value = "Total Traded Volume"
summarytable_header.Cells(6).Value = "Additional Calculations"
summarytable_header.Cells(7).Value = "Ticker"
summarytable_header.Cells(8).Value = "Value"
summarytable_header.Cells(10).Value = "Average Percent Change"
summarytable_header.Cells(11).Value = "% Change representative?"
summarytable_header.Font.Bold = True
summarytable_header.Columns.AutoFit

End Sub

Sub add_services(ws As Worksheet)

'MAIN SCRIPT SETUP
'------------------

Dim sumtablelastrow As Long

Dim sumtable_tickers As Range
Dim sumtable_pct As Range
Dim sumtable_volumes As Range
Dim summmarytable_post As Range

sumtablelastrow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row

Set sumtable_tickers = ws.Range("L2:L" & sumtablelastrow)
Set sumtable_yrly = ws.Range("M2:M" & sumtablelastrow)
Set sumtable_volumes = ws.Range("o2:o" & sumtablelastrow)
Set summarytable_post = Union(sumtable_tickers, sumtable_yrly, sumtable_volumes)

Dim addserv_table As Range
Set addserv_table = ws.Range("Q1:s4")
addserv_table.Cells(2, 1).Value = "Greatest % increase"
addserv_table.Cells(3, 1).Value = "Greatest % decrease"
addserv_table.Cells(4, 1).Value = "Greatest total volume"

'calculation setup

Dim maxIncrease As Double
Dim maxDecrease As Double
Dim maxVolume As Double

Dim maxIncreaseStock As String
Dim maxDecreaseStock As String
Dim maxVolumeStock As String
    
Dim i As Integer

maxIncrease = -1
maxDecrease = 1
maxVolume = -1

'calculations to determin min, max and total volume

For i = 2 To sumtablelastrow
    If sumtable_yrly.Cells(i, 1).Value > maxIncrease Then
        maxIncrease = sumtable_yrly.Cells(i, 1).Value
        maxIncreaseStock = sumtable_tickers.Cells(i, 1).Value
    End If

    If sumtable_yrly.Cells(i, 1).Value < maxDecrease Then
        maxDecrease = sumtable_yrly.Cells(i, 1).Value
        maxDecreaseStock = sumtable_tickers.Cells(i, 1).Value
    End If

    If sumtable_volumes.Cells(i, 1).Value > maxVolume Then
        maxVolume = sumtable_volumes.Cells(i, 1).Value
        maxVolumeStock = sumtable_tickers.Cells(i, 1).Value
    End If
Next i

' PRINTING RESULTS TO ADSERV SUMTABLE
addserv_table.Cells(2, 2).Value = maxIncreaseStock
addserv_table.Cells(3, 2).Value = maxDecreaseStock

addserv_table.Cells(2, 3).Value = (maxIncrease) / 10000
addserv_table.Cells(2, 3).NumberFormat = "0.00%;;"


addserv_table.Cells(3, 3).Value = (maxDecrease) / 10000
addserv_table.Cells(3, 3).NumberFormat = "0.00%"


addserv_table.Cells(4, 2).Value = maxVolumeStock
addserv_table.Cells(4, 3).Value = Format(maxVolume, "#,##0")
addserv_table.Columns.AutoFit


End Sub

