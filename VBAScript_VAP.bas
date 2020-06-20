Attribute VB_Name = "Module1"
Sub multi_year_stock()

' Set variable for ticker name
    Dim tickername As String

' Set variable for total volume per ticker name
    Dim totalvolume As Double
    totalvolume = 0

' Set variable for yearly change
    Dim yrchange As Double
    Dim open_value As Double
    Dim close_value As Double
    
' Set variable for percent change
    Dim percent_change As Double

' Set location for the summary table
    Dim sumtable As Long
    sumtable = 2
    
' Set variable for row index of ticker open value
    Dim open_start As Long
    open_start = 2

 ' Set location for headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yealry Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Change"

' Find the last row
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row

' Loop through ticker volume
    For i = 2 To lastrow

' Check that tickername does not equal same tickername
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
' Set ticker name
    tickername = Cells(i, 1).Value
    
' Add to the total volume
    totalvolume = totalvolume + Cells(i, 7).Value
    
' Calculate the yearly change
    open_value = Cells(open_start, 3).Value
    close_value = Cells(i, 6).Value
    yrchange = close_value - open_value
    
' Caluculate the percent change
    percent_change = Round((yrchange / open_value) * 100, 2)
    On Error Resume Next
    
' Place ticker name in the summary table
    Cells(sumtable, 9).Value = tickername

' Place yearly change in the summary table
    Cells(sumtable, 10).Value = yrchange
    
' Place percent change in the summary table
    Cells(sumtable, 11).Value = "%" & percent_change

' Place total volume in the summary table
    Cells(sumtable, 12).Value = totalvolume
    
' Interior formatting
    If Cells(sumtable, 10).Value > 0 Then
        Cells(sumtable, 10).Interior.ColorIndex = 4
    ElseIf Cells(sumtable, 10).Value < 0 Then
        Cells(sumtable, 10).Interior.ColorIndex = 3
   End If

' Add one to the summary table row
    sumtable = sumtable + 1

' Locate the start of next stop ticker
    open_start = i + 1
    
' Reset total volume
    totalvolume = 0

' If cell following is the same
Else

' Add to the total volume
    totalvolume = totalvolume + Cells(i, 7).Value
    
End If

    Next i

' Set location for challenge headers
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
' Set variables for challenge values
    Dim sheet3, sheet2, sheet1 As Worksheet
    Set sheet3 = Worksheets("2014")
    Set sheet2 = Worksheets("2015")
    Set sheet1 = Worksheets("2016")
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_volume As Double
    
' Establish Max Percent
    max_percent = Application.WorksheetFunction.Max(sheet3.Columns("K"))
    Cells(2, 17).Value = max_percent
   
 ' Establish Min Percent
    min_percent = Application.WorksheetFunction.Min(sheet3.Columns("K"))
    Cells(3, 17).Value = min_percent
    
 ' Establish Max Total Volume
    max_volume = Application.WorksheetFunction.Max(sheet3.Columns("L"))
    Cells(4, 17).Value = max_volume
    
' Match ticker with value
    match_ticker = WorksheetFunction.Match(tickername, sheet3.Range("I2:I2836"), 0)
    sheet3.Range("P2:P4") = match_ticker
   
' Autofit to display data
    Columns("A:R").AutoFit
    

End Sub

