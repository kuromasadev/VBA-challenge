# VBA-challenge
 RICE_RDA_AlbertoPonce

### Stock Analysis 

Prepared in the following framework 

- Stock Counter 
  - Counts stock and prepares summary table with the following 
    - Ticker Symbol
    - Yearly Change of price 
    - Percent Change between open and closing price
  - Matching verification using average stock prices to the percent change, results match if within 0.05% 
- Formatting 
  - Color matching for positive/negative values in the yearly change column 
  - Bold font for summary table headers
  - date value transformed to short date format
- Additional Services 
  - Greatest % Increase
  - Greatest % Decrease
  - Greatest Total Volume



### Main Module reference

> "alpha code", including attempt to run the code across all worksheets 

```visual basic
Sub runthejewels()

Dim wb As Workbook
Set wb = ActiveWorkbook
Dim ws As Worksheet
        
'Establishing for loop to add worksheets to a collection

Dim vba_wallstreet As New Collection
    For Each ws In wb.Worksheets
        vba_wallstreet.Add ws
    Next ws

    'Looping through the developed modules
    
  	For i = 1 To vba_wallstreet.Count
        Set ws = vba_wallstreet(i)
        Call Summary.stockcounter(ws) 'this module performs the stock analysis
        Call Formatting.tableformatting(ws) 'this module formats the table
        Call Formatting.add_services(ws) 'this module does the bonus challenge
    Next i

End Sub
```

## Screenshots

> Note : all screenshots have a cell with the formula =TODAY() as a way to date them

![2018](C:\Users\17138\Dropbox\GRADRICE - DA\RDA_reference_repo\VBA-challenge\stock_2018_screenshot.JPG)

![2019](C:\Users\17138\Dropbox\GRADRICE - DA\RDA_reference_repo\VBA-challenge\stock_2019_screenshot.JPG)

![2020](C:\Users\17138\Dropbox\GRADRICE - DA\RDA_reference_repo\VBA-challenge\stock_2020_screenshot.JPG)

## VBA Script Modules

### Stock Counter

```visual basic
Sub stockcounter(ws As Worksheet)

' INITIAL FORMATTING AND SCRIPT SETUP
'-------------------------
Dim lastrow As Long
lastrow = Cells(ws.Rows.Count, 1).End(xlUp).Row

Dim stocktrades As Range
    Set stocktrades = ws.Range("A2:G" & lastrow)

Dim uniquestocks As New Collection
Dim cell As Range
    
Dim tradedate As Date
    
    'Setting up the list of tickers to track the different columns
    For Each cell In stocktrades.Columns(1).Cells
        Dim ticker As String
        ticker = Trim(cell.Value) ' added trim just in case theres spaces next to the tickers within the data
        If ticker <> "" Then
        'adding tickers to the collection
        'researched language to skip over errors, which I had to add because it would get stuck on labels it had already collected
            On Error Resume Next
            uniquestocks.Add CStr(ticker), CStr(ticker) 'online docs show that CString is a char object
            On Error GoTo 0
        End If
    Next cell

    'formatting the date column so that the summary can identify the dates as actual dates and not divide by 0
    For Each cell In stocktrades.Columns(2).Cells
        If Len(cell.Value) = 8 And IsNumeric(cell.Value) Then
             tradedate = CDate(Left(cell.Value, 4) & "-" & Mid(cell.Value, 5, 2) & "-" & Right(cell.Value, 2))
             cell.Value = tradedate
             cell.NumberFormat = "yyyy-mm-dd"
        End If
    Next cell
               
    'Setting up the summary table consistently starting on column L
    'including a header with the +1
    'this made more sense than to figure out where the cell(x,y) coordinates were.
        
Dim summarytable As Range
Set summarytable = ws.Range("L1:R" & uniquestocks.Count + 1)

'MAIN CALCULATIONS SECTION
'-------------------------

'counter calculator using various worksheetfunctions' taking stock volumnes from column G and using sumif to add them
Dim i As Integer
i = 2

Dim tickername As Variant 'this variant classification seemed to work instead instead of string
    For Each tickername In uniquestocks
        Dim totalvolume As Double
        totalvolume = Application.WorksheetFunction.SumIf(stocktrades.Columns(1), tickername, stocktrades.Columns(7))
        
        'ended up doing something different at first, but this iteration prints the average start price vs average end price as percentage
        'I know it wasn't really asked for, but in terms of representing the entire dataset, by only choosing
        'the first and last trade value wouldn't inform the data consumer of the context behind the change
        ' it could also help compare against similar tickers
        
        Dim valuestock_opensum As Double
        Dim valuestock_closesum As Double
        Dim valuestock_openavg As Double
        Dim valuestock_closeavg As Double
        
        valuestock_opensum = Application.WorksheetFunction.SumIf(stocktrades.Columns(1), tickername, stocktrades.Columns(3))
        valuestock_closesum = Application.WorksheetFunction.SumIf(stocktrades.Columns(1), tickername, stocktrades.Columns(6))
        valuestock_openavg = valuestock_opensum / uniquestocks.Count
        valuestock_closeavg = valuestock_closesum / uniquestocks.Count
        
        Dim avgpct_change As Long
        avgpct_change = ((valuestock_closeavg - valuestock_openavg / valuestock_openavg))
        
                        
        'placing total values in the summary table
        summarytable.Cells(i, 1).Value = tickername
        
        summarytable.Cells(i, 4).Value = totalvolume
        
        summarytable.Cells(i, 10).Value = avgpct_change / 10000
        summarytable.Cells(i, 10).NumberFormat = "0.00%;;;"
        
        'Open - Close Ticker Pricing, and yearly change
        'this one was hard to setup, but was able to place a proxy variable, "last_date_traded" to compare against
        ' I understand it might be extra steps but I think I got it sorted!
                
        Dim valuestock_firstopen As Double
        Dim valuestock_lastclose As Double
        
        Dim yr_start As Date
        Dim yr_end As Date
        Dim last_date_traded As Date
    
        'putting the date formatting at the top also helped with this step
        yr_start = DateSerial(2018, 1, 1)
        yr_end = DateSerial(2020, 12, 31)
        last_date_traded = #1/1/1900#
        
        For Each Row In stocktrades.Rows
            If Row.Columns(1).Value = tickername And Row.Columns(2).Value >= yr_start Then
                valuestock_firstopen = Row.Columns(3).Value
                Exit For                                                   ' Exit loop once we find the first matching row
            End If
        Next Row
            
        For Each Row In stocktrades.Rows
            If Row.Columns(1).Value = tickername And Row.Columns(2).Value <= yr_end And Row.Columns(3).Value > last_date_traded Then
                last_date_traded = Row.Columns(3).Value
                valuestock_lastclose = Row.Columns(6).Value                ' Exit loop once we find the first matching row
                Exit For
            End If
        Next Row
                       
         'take open price as the value closes to the start of the year, and the close value closes to the end of the year
        Dim pct_change As Double
        Dim yrly_change As Double
        
        pct_change = (valuestock_lastclose - valuestock_firstopen) / valuestock_firstopen
        yrly_change = (valuestock_lastclose - valuestock_firstopen)
    
        summarytable.Cells(i, 2).Value = yrly_change
        summarytable.Cells(i, 3).Value = pct_change
    
      
        i = i + 1
    Next tickername

End Sub

```



### Table Formatter

```visual basic
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
```



### Additional Services (Bonus)

```visual basic
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
                ' this section was attempted with ranges but it didn't work

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
            'for whatever reason, the percentages were getting having the decimal places moved by 4 to the right, so my rudimentary solution was to move them back in place by dividing by 10,000 and then apply the percent format
            
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
```

