Attribute VB_Name = "Module1"
Sub tickerTracker()
     For Each ws In Worksheets
        '****************************
        '
        '****************************
        Dim wsRows As Long
        Dim wsCol As Integer
        

        '****************************
        'Define Summary Table Cell Variables
        '****************************
        Dim sumRow As Integer
        Dim sumTickerCol As Integer
        Dim sumDiffCol As Integer
        Dim sumPercentCol As Integer
        Dim sumVolumeCol As Integer
        
        
        '****************************
        'Identify Summary Table Cells
        '****************************
        sumRow = 2
        sumTickerCol = 10
        sumDiffCol = 11
        sumPercentCol = 12
        sumVolumeCol = 13
        
        '****************************
        ' Create Summary Table Variables
        '****************************
        Dim tickerName As String
        Dim firstOpen As Double
        Dim lastClose As Double
        Dim summDiff As Double
        Dim sumPerc As Double
        Dim currVol As LongLong
        
        '****************************
        'Identify last row and column of data
        '****************************
        wsRows = ws.Cells(1, 1).End(xlDown).Row
        wsCol = ws.Cells(1, 1).End(xlToRight).Column
        
        '****************************
        'Create Headers for Summary Table
        '****************************
        ws.Cells(1, sumTickerCol).Value = "Ticker"
        ws.Cells(1, sumDiffCol).Value = "Difference"
        ws.Cells(1, sumPercentCol).Value = "Percent Change"
        ws.Cells(1, sumVolumeCol).Value = "Volume"
        
        '****************************
        'Initial Ticker Name and Opening Variables
        '****************************
        tickerName = ws.Cells(2, 1).Value
        firstOpen = ws.Cells(2, 3).Value
        'msgBox tickerName
        'msgBox firstOpen
        
        '****************************
        'Loop through rows
        '****************************
        For r = 2 To wsRows
            If ws.Cells(r, 1) <> ws.Cells(r + 1, 1) Then
                lastClose = ws.Cells(r, 6).Value
                'msgBox lastClose
                ws.Cells(sumRow, sumTickerCol).Value = tickerName
                summDiff = lastClose - firstOpen
                'msgBox summDiff
                
                ws.Cells(sumRow, sumDiffCol).Value = summDiff
                If ws.Cells(sumRow, sumDiffCol).Value >= 0 Then
                    ws.Cells(sumRow, sumDiffCol).Interior.Color = RGB(0, 128, 0)

                ElseIf ws.Cells(sumRow, sumDiffCol).Value < 0 Then
                    ws.Cells(sumRow, sumDiffCol).Interior.Color = RGB(128, 0, 0)

                End If
                ws.Cells(sumRow, sumDiffCol).NumberFormat = "$0.00"
                
                If firstOpen <> 0 Then
                    sumPerc = (lastClose - firstOpen) / firstOpen
                Else
                    sumPerc = 0
                End If

                ws.Cells(sumRow, sumPercentCol).Value = sumPerc
                ws.Cells(sumRow, sumPercentCol).NumberFormat = "0.00%"

                ws.Cells(sumRow, sumVolumeCol).Value = currVol
                
                
                sumRow = sumRow + 1
                
                tickerName = ws.Cells(r + 1, 1).Value
                firstOpen = ws.Cells(r + 1, 3).Value
                currVol = 0
                'msgBox currVol
                
            Else
                currVol = currVol + ws.Cells(r, 7).Value
                'msgBox currVol

            End If
        Next r

    Next ws
End Sub

