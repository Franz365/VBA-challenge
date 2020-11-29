Sub Multiple_year_stock_analysis():

    'Define variables
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim ticker As String
    Dim OpenV As Double
    Dim CloseV As Double
    Dim YearlyChange As Double
    Dim PercentChange As Variant
    Dim StockVol As Double
    Dim SummaryTable As Range
    
    'Start loop through all worksheets
    For Each ws In Worksheets
    
        'Select worksheet (and then the next worksheet)
        Worksheets(ws.Name).Select
    
        'Determine the last row (Ctrl + Shift + End)
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'Determine the last column (Ctrl + Shift + End)
        LastColumn = ws.Cells(7, ws.Columns.Count).End(xlToLeft).Column
        
        'Write summary table headers
        Cells(1, LastColumn + 2) = "#"
        Cells(1, LastColumn + 3) = "Ticker"
        Cells(1, LastColumn + 4) = "Yearly Change"
        Cells(1, LastColumn + 5) = "Percent Change"
        Cells(1, LastColumn + 6) = "Total Stock Volume"
        
        ' Keep track of the location for each credit card brand in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        'Loop through all stocks
        For i = 2 To LastRow
    
            ' Check if we are still within the ticker, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Ticker Count #
                ' Print the ticker count in the Summary Table
                Cells(Summary_Table_Row, LastColumn + 2).Value = Summary_Table_Row - 1
                
                'Ticker
                ' Set the ticker
                ticker = Cells(i, 1).Value
                ' Print the ticker in the Summary Table
                Cells(Summary_Table_Row, LastColumn + 3).Value = ticker
                
                'Yearly Change
                'set the closing price
                CloseV = Cells(i, LastColumn - 1).Value
                'calculate the Yearly Change
                YearlyChange = CloseV - OpenV
                'print the Yearly change in the summary table
                Cells(Summary_Table_Row, LastColumn + 4) = YearlyChange
                'change cell formatting
                Cells(Summary_Table_Row, LastColumn + 4).Style = "Currency"
                'Format negative values red, positive values green, 0 values stay white
                If Cells(Summary_Table_Row, LastColumn + 4).Value < 0 Then
                    Cells(Summary_Table_Row, LastColumn + 4).Interior.ColorIndex = 3
                ElseIf Cells(Summary_Table_Row, LastColumn + 4).Value > 0 Then
                    Cells(Summary_Table_Row, LastColumn + 4).Interior.ColorIndex = 4
                End If
                            
                'Percent Change
                'If clause to avoide 0 division error
                If OpenV = 0 Then
                    'If division is not possible, print "NA" in the summary table
                    Cells(Summary_Table_Row, LastColumn + 5) = "NA"
                Else
                    'calculate the percentage change
                    PercentChange = YearlyChange / OpenV
                    'print the Percent Change in the summary table
                    Cells(Summary_Table_Row, LastColumn + 5) = PercentChange
                    'Cell format as Percent
                    Cells(Summary_Table_Row, LastColumn + 5).NumberFormat = "0.00%"
                End If
                                                                    
                'Total Stock Volume
                ' Add to the Stock Volume
                StockVol = StockVol + Cells(i, LastColumn).Value
                ' Print the stock volume in the Summary Table
                Cells(Summary_Table_Row, LastColumn + 6).Value = StockVol
                Cells(Summary_Table_Row, LastColumn + 6).Style = "Comma [0]"
              
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset the stock volume
                StockVol = 0
                
                'Reset the openv
                OpenV = 0
                
            'If the cell immidiately following a row is the same brand...
            Else
            
                'Add to the stock volume
                StockVol = StockVol + Cells(i, LastColumn).Value
                
                'Store stock open price in variable
                If OpenV = 0 Then
                
                    OpenV = Cells(i, LastColumn - 4).Value
                
                End If
                            
            End If
        
        Next i
        
        'AutoFit Summary table
        Columns("I:M").AutoFit
    
    Next ws

End Sub

