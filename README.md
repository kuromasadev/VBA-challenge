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

> attempt to run the code across all worksheets 

```visual basic
Sub runthejewels()

Dim wb As Workbook
Set wb = ActiveWorkbook
Dim ws As Worksheet

Dim vba_wallstreet As New Collection
    For Each ws In wb.Worksheets
        vba_wallstreet.Add ws
    Next ws

  For i = 1 To vba_wallstreet.Count
        Set ws = vba_wallstreet(i)
        Call Summary.stockcounter(ws)
        Call Formatting.tableformatting(ws)
        Call Formatting.add_services(ws)
    Next i

End Sub
```



