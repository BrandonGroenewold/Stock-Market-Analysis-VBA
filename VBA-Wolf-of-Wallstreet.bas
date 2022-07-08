Attribute VB_Name = "Module1"
Sub StockData()
'Define the variables I am working with
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    Dim ws As Worksheet
    Dim SummaryTable As Integer
    Dim VolumeRow As Long
    
    
    
'Have code go through each worksheet
For Each ws In ThisWorkbook.Worksheets


'Define column names for new table


    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'Google says setup integers for loop
    SummaryTable = 2
    VolumeRow = 2
    'Formula to count how many rows there are
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
'Start loop
    'Start on row 2 and go to last row
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'find variables
            Ticker = Cells(i, 1).Value
            'Volume = Cells(i, 7).Value
            YearOpen = Cells(VolumeRow, 3).Value
            YearClose = Cells(i, 6).Value
            YearlyChange = YearClose - YearOpen
            PercentChange = (YearClose - YearOpen) / YearOpen
            
            Volume = Application.WorksheetFunction.Sum(Range("G" & VolumeRow & ":" & "G" & i))
            VolumeRow = i + 1
        'Special thanks to Sanoo Singh for helping me with the line of code above. I couldn't have figure it out without him
            
            'Putting variables found into the new table
                ws.Cells(SummaryTable, 9).Value = Ticker
                ws.Cells(SummaryTable, 10).Value = YearlyChange
                ws.Cells(SummaryTable, 11).Value = PercentChange
                ws.Cells(SummaryTable, 12).Value = Volume
                
                
 'Format percent change column to percentage
    ws.Cells(SummaryTable, 11).NumberFormat = "0.00%"
       
       'Conditional formatting for Yearly Change
       
             If ws.Cells(SummaryTable, 10).Value >= 0 Then
                ws.Cells(SummaryTable, 10).Interior.ColorIndex = 4
            Else
              ws.Cells(SummaryTable, 10).Interior.ColorIndex = 3
            End If
       SummaryTable = SummaryTable + 1
    'Make sure to change it to conditional formating instead of just a formula checking it once it is ran
   
    
        End If
        Next i
Next ws

End Sub

