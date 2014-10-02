Sub PrintColumns()
    ' Print-Columns - Excel macro
   
    ' Print-Columns allows to and quickly easily fit sheets with a 
    ' high number of rows and low
    ' number of columns on a single page. This helps show more data on
    ' each papersheet and save paper.
     
    ' HOW IT WORKS:
     
    ' Divides selected cells in blocks of height h, then creates a
    ' new sheet. It then puts each block side by side, up to
    ' PAGECOLS blocks. This is one page.
    ' It then proceeds to create a new page under the first one.
   
    ' Both h and pagecols are set by the user when run.
   
    ' It also automatically sets page width to one page and
    ' adds page breaks after eery new page, for a print friendly format.
 
    ' Important: all selected cells should have the same height to avoid
    ' weird impagination
   
    ' This software is placed in the public domain by its creator.   

    Dim coln, colstart, rown, rowstart, h, n, rowi, coli, pagecols, pageold As Integer
    coln = Selection.Columns.Count
    rown = Selection.rows.Count
    rowstart = Selection.Cells(1).Row
    colstart = Selection.Cells(1).Column
    pageold = 0
   
    h = 80          ' number of rows to divide the sheet at
    pagecols = 2    ' number of broken columns to insert in each page
   
    h = Application.InputBox("Enter number of lines for each page", Type:=1)
    pagecols = Application.InputBox("Enter number of columns for each page", Type:=1)
   
    Set wsold = ActiveSheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
   
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
 
    For i = 1 To Int(WorksheetFunction.Ceiling_Precise(CDbl(rown / h), 1))
   
        Dim p1x, p1y, p2x, p2y As Integer
        p1x = rowstart + h * (i - 1)
        p1y = colstart
        p2x = p1x + h - 1
        p2y = p1y + coln - 1
        remaining = rown - h * (i - 1) - 1
 
        If remaining < h Then
            p2x = p1x + remaining
        Else
            remaining = h
        End If
       
        Dim r1x, r1y, r2x, r2y, pagenum As Integer
        pagenum = Int(WorksheetFunction.Ceiling_Precise(CDbl(i / pagecols), 1))
        r1x = (pagenum - 1) * (h + 1) + 1
        r1y = (coln + 1) * (i - 1 - (pagenum - 1) * pagecols) + 1
        r2x = r1x + remaining - 1
        r2y = r1y + coln - 1
       
        wsold.Range(wsold.Cells(p1x, p1y), wsold.Cells(p2x, p2y)).Copy Destination:=ws.Range(ws.Cells(r1x, r1y), ws.Cells(r2x, r2y))
       
        If pagenum = 1 Then
            For j = 1 To coln
                ws.Cells(1, r1y + j - 1).ColumnWidth = wsold.Cells(rowstart, colstart + j - 1).ColumnWidth
            Next j
        End If
 
        If pageold <> pagenum Then  ' new page
           If pagenum <> 1 Then
                ws.HPageBreaks.Add Before:=ws.Cells(r1x, 1)
            End If
        ' Uncomment to set row height too (much slower, not reccomended):
       ' For j = 1 To h
       '     ws.Cells(r1x + j - 1, 1).RowHeight = wsold.Cells(rowstart + j - 1, colstart).RowHeight
       ' Next j
       End If
       
        pageold = pagenum
    Next i
End Sub