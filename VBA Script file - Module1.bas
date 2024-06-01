Attribute VB_Name = "Module1"
Sub tickerchallenge()

' Set initial variable types
Dim TickerName As String
Dim LastRow As Long
Dim TickerCount As Integer
Dim TotalStockVolume As Double
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim QuarterlyChange As Double
Dim PercentChange As Double
Dim ws As Worksheet

' Loop through each worksheet in workbook
For Each ws In Worksheets

' Set initial variables values for new ws
TickerCount = 1
TickerName = ws.Range("A2").Value
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
OpeningPrice = ws.Range("C2").Value

'Print column headings
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Quarterly Change"
ws.Range("K1") = "PercentChange"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"
ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Total Volume"



    ' Loop through all ticker rows to last row
    For i = 2 To LastRow
        
        'Track running total of daily volume per ticker
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
        ' If same ticker name, then
        If ws.Cells(i + 1, 1).Value = TickerName Then
        
        ' If this is last line for ticker then print values to summary table and reset values
        Else
        'Calculate Quarterly and Percent Change
        ClosingPrice = ws.Cells(i, 6).Value
        QuarterlyChange = ClosingPrice - OpeningPrice
        PercentChange = QuarterlyChange / OpeningPrice
        
        ' Print to Summary Table
        ws.Range("I" & TickerCount + 1).Value = TickerName
        ws.Range("J" & TickerCount + 1).Value = QuarterlyChange
        ws.Range("K" & TickerCount + 1).Value = PercentChange
        ws.Range("L" & TickerCount + 1).Value = TotalStockVolume
        
            'Format Summary Table for pos and neg quarterly change
            If QuarterlyChange < 0 Then
            ' Set the Quarterly Change Cell Colours to Red
            ws.Range("J" & TickerCount + 1).Interior.ColorIndex = 3
            Else
            ' Set the Quarterly Change Cell Colour to Green
            ws.Range("J" & TickerCount + 1).Interior.ColorIndex = 4
                If QuarterlyChange = 0 Then
                ' Set the Quarterly Change Cell Colours to White
                ws.Range("J" & TickerCount + 1).Interior.ColorIndex = 2
                End If
            End If
            
        ' Format Summary Table for the Percent Change to 2 decimal places with a % sign
        ws.Range("K" & TickerCount + 1).NumberFormat = "0.00%"
        ' Format Summary Table for the Quarterly Change to 2 decimal places with a % sign
        ws.Range("J" & TickerCount + 1).NumberFormat = "0.00"
        
        'Reset Values for new ticker name
        TickerCount = TickerCount + 1
        TickerName = ws.Cells(i + 1, 1).Value
        OpeningPrice = ws.Cells(i + 1, 3).Value
        ClosingPrice = ws.Cells(i, 6).Value
        TotalStockVolume = 0
    End If

    Next i

'------------------------------------------------------------------------------------------
'Bonus Work - After each ws summary return Ticker and Value
'------------------------------------------------------------------------------------------
Dim GPI_Ticker As String
Dim GPD_Ticker As String
Dim GTotVol_Ticker As String
Dim GPI_Value As Double
Dim GPD_Value As Double
Dim GTotVol_Value As Double

'Set initia; values for variables

GPI_Value = ws.Cells(2, 11).Value
GPD_Value = ws.Cells(2, 11).Value
GTotVol_Value = ws.Cells(2, 12).Value

    For i = 2 To TickerCount
    
        'Look For Biggest % Increase and set related ticker
        If GPI_Value <= ws.Cells(i, 11).Value Then
        GPI_Value = ws.Cells(i, 11).Value
        GPI_Ticker = ws.Cells(i, 9).Value
        End If
        
        'Look For Biggest % Decrease and set related ticker
        If GPD_Value >= ws.Cells(i, 11).Value Then
        GPD_Value = ws.Cells(i, 11).Value
        GPD_Ticker = ws.Cells(i, 9).Value
        End If
        
        'Look For Biggest Total Stock Volume and set related ticker
        If GTotVol_Value <= ws.Cells(i, 12).Value Then
        GTotVol_Value = ws.Cells(i, 12).Value
        GTotVol_Ticker = ws.Cells(i, 9).Value
        End If
        
    Next i


' Print to Summary Table and format
        ws.Range("O" & 2).Value = GPI_Ticker
        ws.Range("O" & 3).Value = GPD_Ticker
        ws.Range("O" & 4).Value = GTotVol_Ticker
        ws.Range("P" & 2).Value = GPI_Value
        ws.Range("P" & 3).Value = GPD_Value
        ws.Range("P" & 4).Value = GTotVol_Value

Next ws

End Sub


