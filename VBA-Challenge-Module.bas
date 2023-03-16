Attribute VB_Name = "Module1"
Sub Ticker()
    
    ' Declare All Variables
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim lastTick As String
    Dim tickerCount As Integer
    
    Dim totalVolume As LongLong
    Dim openValue, closeValue As Double
    Dim percentChange, yearlyChange As Double
    
    Dim greatestIncrease, greatestDecrease As Double
    Dim greatestVolume As LongLong
    Dim greatestIncreaseTicker, greatestDecreaseTicker, greatestVolumeTicker As String
    
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
        ' Changes to the worksheet we are working on
        ws.Select
    
        ' Initialize Variables
        lastTick = ""
        tickerCount = 0
        totalVolume = 0
        percentChange = 1
        yearlyChange = 1
        openValue = 1
        closeValue = 1
        
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Create Array that stores titles for columns
        Dim titleArr(4) As String
        titleArr(0) = "Ticker"
        titleArr(1) = "Yearly Change"
        titleArr(2) = "Percent Change"
        titleArr(3) = "Total Stock Volume"
        
        ' Write Column Titles
        For i = 0 To 3
            Cells(1, i + 9).Value = titleArr(i)
        Next
        
    
        ' Loop through all rows
        For i = 2 To lastrow
        
            If lastTick = Cells(i, 1).Value Then
                totalVolume = totalVolume + Cells(i, "G").Value
                closeValue = Cells(i, "F").Value
            Else
                ' Calculate Data
                yearlyChange = closeValue - openValue
                percentChange = yearlyChange / openValue
                
                ' Prevents from writing on title row
                If tickerCount > 0 Then
                    ' Write Data
                    Cells(tickerCount + 1, "I").Value = lastTick
                    Cells(tickerCount + 1, "J").Value = yearlyChange
                    Cells(tickerCount + 1, "K").Value = percentChange
                    Cells(tickerCount + 1, "L").Value = totalVolume
                    
                    ' Calculate greatestVolume
                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        greatestVolumeTicker = lastTick
                    End If
            
                    ' Change color based on value
                    If yearlyChange < 0 Then
                        Cells(tickerCount + 1, "J").Interior.ColorIndex = 3
                    Else
                        Cells(tickerCount + 1, "J").Interior.ColorIndex = 4
                    End If
                    
                    ' Format Cells
                    Cells(tickerCount + 1, "J").NumberFormat = "0.00"
                    Cells(tickerCount + 1, "K").NumberFormat = "0.00%"
                End If
                
                ' Calculate greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = lastTick
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = lastTick
                End If
                
                lastTick = Cells(i, 1).Value
                tickerCount = tickerCount + 1
                openValue = Cells(i, "C").Value
                totalVolume = Cells(i, "G").Value
                
            End If
            
        Next
        
        ' Print Column Titles
        Cells(1, "O").Value = "Ticker"
        Cells(1, "P").Value = "Value"
        
        ' Print Row Titles
        Cells(2, "N").Value = "Greatest % Increase"
        Cells(3, "N").Value = "Greatest % Decrease"
        Cells(4, "N").Value = "Greatest Total Volume"
        
        ' Print Values
        Cells(2, "O").Value = greatestIncreaseTicker
        Cells(2, "P").Value = greatestIncrease
        
        Cells(3, "O").Value = greatestDecreaseTicker
        Cells(3, "P").Value = greatestDecrease
        
        Cells(4, "O").Value = greatestVolumeTicker
        Cells(4, "P").Value = greatestVolume
        
        ' Format Values
        Cells(2, "P").NumberFormat = "0.00%"
        Cells(3, "P").NumberFormat = "0.00%"
        Cells(4, "P").NumberFormat = "0.00E+00"
        
        ' AutoFit All Cells
        Columns("I:P").AutoFit

    Next ws
End Sub
