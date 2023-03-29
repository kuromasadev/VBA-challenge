Attribute VB_Name = "Alpha_Module"
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
