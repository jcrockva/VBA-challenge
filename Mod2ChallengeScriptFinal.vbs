Attribute VB_Name = "Module1"
Sub stockAnalysis():

    'set variable to hold total stock volume
    Dim total As Double
    
    'Set control variable to loop through rows
    Dim row As Long
    
    'Set variable that holds total number of rows
    Dim rowCount As Double
    
    'Set variable that holds the quarterly change
    Dim quarterlyChange As Double
    
    'Set variable that holds % change
    Dim percentChange As Double
    
    'Set variable that holds rows of the summary table row
    Dim summaryTable As Long
    
    'Set variable that holds start of a stocks rows
    Dim stockStartRow As Long
    
    'Set variable that holds start of a stocks first open
    Dim startValue As Long
    
    'Set variable that locates the last ticker
    Dim lastTicker As String
    
    'loop through all worksheets in workbook
    For Each ws In Worksheets
    
        'Set the title row of the summary section
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Set the title row of the aggregate section
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'initialize values
        'Summary table row starts at 0
        summaryTableRow = 0
        
        'total stock volume starts at 0
        total = 0
        
        'quarterly change starts at 0
        quarterlyChange = 0
        
        'first stock starts on row 2
        stockStartRow = 2
        
        'first open starts on row 2
        startValue = 2
        
        'set value of the last row in sheet
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        'set value for last ticker
        lastTicker = ws.Cells(rowCount, 1).Value
        
        'loop through the sheet
        For row = 2 To rowCount
        
            'check if ticker has changed
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            
                'if there is a change
                
                'add to the total stock volume
                total = total + ws.Cells(row, 7).Value
                
                'check if total stock volume is 0
                If total = 0 Then
                
                    'print results in sumary table
                    
                    'print the ticker
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    'print 0 for quarterly change
                    ws.Range("J" & 2 + summaryTableRow).Value = 0
                    'print 0 for % change
                    ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"
                    'print 0 for total stock volume
                    ws.Range("L" & 2 + summaryTableRow).Value = 0
                Else
                    'find first non 0 first opoen
                    If ws.Cells(startValue, 3).Value = 0 Then
                        'if first open is 0 move to next row
                        For findValue = startValue To row
                        
                            'check next or rows after
                            If ws.Cells(findValue, 3).Value <> 0 Then
                               'update start value
                                startValue = findValue
                                'leave loop
                                Exit For
                            End If
                        
                        Next findValue
                    End If
                    
                    'calculate the quarterly change
                    quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(startValue, 3).Value
                    
                    'calculate percent change
                    percentChange = quarterlyChange / ws.Cells(startValue, 3).Value
                    
                    'print results
                    'print the ticker
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    'print quarterly change
                    ws.Range("J" & 2 + summaryTableRow).Value = quarterlyChange
                    'print % change
                    ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                    'print total stock volume
                    ws.Range("L" & 2 + summaryTableRow).Value = total
                    
                    'format color changes for quarterly change column in summary section
                    If quarterlyChange > 0 Then
                        'color green
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
                    ElseIf quarterlyChange < 0 Then
                        'color red
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
                    Else
                        'no change to cell
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
                    End If
                    
                    'reset the values for next ticker
                    'reset total stock volume
                    total = 0
                    'resert average change
                    averageChange = 0
                    'reset quarterly change
                    quarterlyChange = 0
                    'move start row next row in sheet
                    startValue = row + 1
                    'move to next row in summary table
                    summaryTableRow = summaryTableRow + 1
                    
                End If
                    
            Else
                'if there is no ticker change keep adding to total volume
                'get value from the 7th column
                total = total + ws.Cells(row, 7).Value
                       
            End If
               
        Next row
        
        'prevent extra data in summary section
        'find last row of data in summary table
            
        'update summary table row
        summaryTableRow = ws.Cells(Rows.Count, "I").End(xlUp).row
            
        'find the lasst data from extra rows J-L
        Dim lastExtraRow As Long
        lastExtraRow = ws.Cells(Rows.Count, "J").End(xlUp).row
            
        'create loop that clears extra data
        For e = summaryTableRow To lastExtraRow
            'for loop for 9-12
            For Column = 9 To 12
                ws.Cells(e, Column).Value = ""
                ws.Cells(e, Column).Interior.ColorIndex = 0
            Next Column
        Next e
            
        'print the summary aggregates
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2))
        
        'Find ticker names using match function
        Dim greatestIncreaseRow As Double
        Dim greatestDecreaseRow As Double
        Dim greatestTotalVolRow As Double
        greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
        greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
        greatestTolVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2)), ws.Range("L2:L" & summaryTableRow + 2), 0)
    
        'display ticker symbols for Aggragate Table
        ws.Range("P2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
        ws.Range("P3").Value = ws.Cells(greatestDecreaseRow + 1, 9).Value
        ws.Range("P4").Value = ws.Cells(greatestTolVolRow + 1, 9).Value
        
        'format the summary table columns
        For s = 0 To summaryTableRow
            ws.Range("J" & 2 + s).NumberFormat = "0.00"
            ws.Range("K" & 2 + s).NumberFormat = "0.00%"
            ws.Range("L" & 2 + s).NumberFormat = "#,###"
        Next s
        
        'format the summary aggregates
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,###"
        
        'Autofit info across all columns
        ws.Columns("A:Q").AutoFit
        
    Next ws
    
End Sub

