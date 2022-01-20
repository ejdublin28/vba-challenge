Attribute VB_Name = "Module1"
Sub LoopWorksheet()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call AnnualStockSummary
    Next
    Application.ScreenUpdating = True
End Sub

Sub AnnualStockSummary()
    
    Dim j As Integer
    Dim lastrow As Long
    Dim ticker As Variant
   
    Dim min_index As Variant
    Dim max_index As Long
    
    Dim clos As Double
    Dim opn As Double
    
    Dim min_date As Long
    Dim max_date As Long
    
    'Pulls list of unique tickers
    Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True
    
    'Set Up Column Headers
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Establish variables within sheet
    min_date = WorksheetFunction.Min(Range("B:B"))
    max_date = WorksheetFunction.Max(Range("B:B"))
    
    lastrow = Cells(Rows.Count, 9).End(xlUp).row
    lastrow_full = Cells(Rows.Count, 1).End(xlUp).row

    
    'Update information for each ticker
    For j = 2 To lastrow
        ticker = Cells(j, 9).Value
        
        For row = 2 To lastrow_full
            If Cells(row, 2).Value = min_date And Cells(row, 1) = ticker Then
                opn = Cells(row, 3).Value
            End If
            
            If Cells(row, 2).Value = max_date And Cells(row, 1) = ticker Then
                clos = Cells(row, 6).Value
            End If
        Next row
       
        'Saving YearlyChange, % Change and Total Stock Volume
        Cells(j, 10).Value = (clos - opn)
        Cells(j, 11).Value = (clos - opn) / opn
        Cells(j, 11).NumberFormat = "0.00%"
        Cells(j, 12).Value = WorksheetFunction.SumIf(Range("A:A"), ticker, Range("G:G"))
        
        'Format Cell Color Based On Value
        If Cells(j, 10).Value >= 0 Then
            Cells(j, 10).Interior.Color = vbGreen
        Else
            Cells(j, 10).Interior.Color = vbRed
        End If
        
    Next j
    
End Sub
