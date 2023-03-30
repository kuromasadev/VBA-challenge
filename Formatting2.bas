Attribute VB_Name = "Formatting2"
Sub missingformat()

Dim summarytable_header As Range
Set summarytable_header = Range("L1:Z1")

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
