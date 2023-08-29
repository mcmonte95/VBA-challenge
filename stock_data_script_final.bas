Attribute VB_Name = "Module1"

Sub AnalyzeStocks()

'Define variables with type
Dim finalRow As Long
Dim finalColumn As Long
Dim yearOpen As Double
Dim yearClose As Double
Dim cell As Range
Dim count As Long
Dim vol As Double
Dim g_increase As Double
Dim g_decrease As Double
Dim g_volume As Double
Dim name As String

'Iterate through all sheets using 'For Each" loop
For Each ws In Worksheets

    'Initialize variables for a new sheet
    name = ws.name
    vol = 0
    count = 2
    finalRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    finalColumn = 7
    
    'Define results table for the sheet
    ws.Cells(1, finalColumn + 2).Value = "Ticker"
    ws.Cells(1, finalColumn + 3).Value = "Yearly Change"
    ws.Cells(1, finalColumn + 4).Value = "Percent Change"
    ws.Cells(1, finalColumn + 5).Value = "Total Stock Volume"
    
    '(1) Use 'for each' loop to act on the Range
    For Each cell In ws.Range("A2:A" & finalRow)
            
                     
        'Set stock opening price for the year
        If cell.Offset(0, 1).Value = name & "0102" Then
        
           yearOpen = cell.Offset(0, 2).Value
        
        End If
    
        'increment total stock volume
        vol = vol + cell.Offset(0, 6).Value
    
    
        'Start populating results table by checking current cell and one below to see if they match
        If cell <> cell.Offset(1, 0).Value Then
            
            'Get unique ticker symbols and put in "Ticker" column of results table
            ws.Cells(count, finalColumn + 2).Value = cell.Value
            
            'Set stock closing price for the year
            yearClose = cell.Offset(0, 5).Value
            
            'Set "Yearly Change" as yearClose - yearOpen (price delta) and apply formatting for gain and loss
            ws.Cells(count, finalColumn + 3).Value = yearClose - yearOpen
            If ws.Cells(count, finalColumn + 3).Value > 0 Then
                ws.Cells(count, finalColumn + 3).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(count, finalColumn + 3).Value < 0 Then
                ws.Cells(count, finalColumn + 3).Interior.ColorIndex = 3
            
            End If
            
            '(2) Set "Percent Change" using formula (yearClose - yearOpen)/yearOpen and format as a percentage (percent change)
            ws.Cells(count, finalColumn + 4).Value = (yearClose - yearOpen) / yearOpen
            ws.Cells(count, finalColumn + 4).NumberFormat = "0.00%"
            
            'Set "Total Stock Volume" to vol which has the sum of the volume for each day of that stock ticker
            ws.Cells(count, finalColumn + 5).Value = vol
            ws.Cells(count, finalColumn + 5).NumberFormat = "0"
            
            'increment count for results table and reset vol to 0
            count = count + 1
            vol = 0
            
        End If

    Next
    
'-----------------------------------------------------------------------------------------------------------------------
    
    'Re-initialize variables so we can iterate through the range of the newly generated table
    finalRow = ws.Cells(ws.Rows.count, finalColumn + 2).End(xlUp).Row
    finalColumn = 12
    Set cell = Nothing
    g_increase = 0
    g_decrease = 0
    g_volume = 0
    
    'Define new table for "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
    ws.Cells(2, finalColumn + 3).Value = "Greatest % increase"
    ws.Cells(3, finalColumn + 3).Value = "Greatest % Decrease"
    ws.Cells(4, finalColumn + 3).Value = "Greatest Total Volume"
        
    ws.Cells(1, finalColumn + 4).Value = "Ticker"
    ws.Cells(1, finalColumn + 5).Value = "Value"
    
    'Iterate through results table
    For Each cell In ws.Range("I2:I" & finalRow)
    
        'Find "Greatest % Increase"
        If cell.Offset(0, 2).Value > g_increase Then
            
            g_increase = cell.Offset(0, 2).Value
            ws.Cells(2, finalColumn + 4).Value = cell.Value
            ws.Cells(2, finalColumn + 5).Value = g_increase
            ws.Cells(2, finalColumn + 5).NumberFormat = "0.00%"
            
        End If
    
        'Find "Greatest % Decrease"
        If cell.Offset(0, 2).Value < g_decrease Then
            
            g_decrease = cell.Offset(0, 2).Value
            ws.Cells(3, finalColumn + 4).Value = cell.Value
            ws.Cells(3, finalColumn + 5).Value = g_decrease
            ws.Cells(3, finalColumn + 5).NumberFormat = "0.00%"
            
        End If
    
        'Find "Greatest Total Volume"
        If cell.Offset(0, 3).Value > g_volume Then
            
            g_volume = cell.Offset(0, 3).Value
            ws.Cells(4, finalColumn + 4).Value = cell.Value
            ws.Cells(4, finalColumn + 5).Value = g_volume
            ws.Cells(4, finalColumn + 5).NumberFormat = "0"
            
        End If
        
    
    Next
    
    
Next

End Sub
