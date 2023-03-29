Attribute VB_Name = "Summary"
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

