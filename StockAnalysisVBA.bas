Attribute VB_Name = "Module1"
Sub StocksDataAnalysis()

    ' We're creating a Macro to analyze stock data in each worksheet of this file in order to gain insight into
    ' changes in price and volume across different quarters.
    
    ' We declare a worksheet variable first
    
    Dim ws As Worksheet
    
    ' Declare all variables needed to run our Macro.
    
    Dim Ticker As String
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String
     
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TotalVolume As Double
    
    Dim QurterlyChange As Double
    Dim PercentChange As Double
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    
    Dim SummaryRow As Integer
    
    Dim LastRow As Long
    Dim i As Long
    
    ' Loop through every worksheet
    
    For Each ws In Worksheets
    
        ' Switch and activate this sheet.
        
        ws.Activate
        
        ' Define the initial values before running our Macro.
        
        SummaryRow = 2
        TotalVolume = 2
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0
            
        ' Define last row with data.
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Create headers for summary table.
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quaterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Begin a For Loop through all stock data.
        
        For i = 2 To LastRow
        
            ' Locate first time ticker appears.
            
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ' Take opening price from column #.3
            
            OpenPrice = ws.Cells(i, 3).Value
                    
             End If
         
            ' Add the volume for from colum 7.
        
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' If the ticker name changes, then choose the closing price on column 6 for that ticker.
                
                Ticker = ws.Cells(i, 1).Value
                
                ClosePrice = ws.Cells(i, 6).Value
        
                ' Calculate the quaterly change
                
                QuaterlyChange = ClosePrice - OpenPrice
                
                 ' Calculate percentage change.
                 
                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = QuaterlyChange / OpenPrice
                End If
                
                ' Write data to the summary table.
                
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = QuaterlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                ' Format the Percent Change column.
                
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                
                ' Color the quaterly change column based on the outcome (positive or negative).
        
                If ws.Cells(SummaryRow, 11).Value > 0 Then
                    ws.Cells(SummaryRow, 11).Interior.Color = vbGreen
                Else
                    ws.Cells(SummaryRow, 11).Interior.Color = vbRed
                End If
                
                ' Check if the ticker has the greatest volume.
                
                If TotalVolume > MaxVolume Then
                    MaxVolume = TotalVolume
                    MaxVolumeTicker = Ticker
                End If
                
                ' Continue onto the next row in the summary table.
                
                SummaryRow = SummaryRow + 1
                
                ' Reset total volume for the next ticker.
                
                TotalVolume = 0
                
            End If
            
        Next i
    
        ' Output the greatest values after going through all tickers in the worksheet.
    
        ws.Cells(2, 15).Value = "Ticker"
        ws.Cells(2, 16).Value = "Value"
        
        ws.Cells(3, 14).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = MaxIncreaseTicker
        ws.Cells(3, 16).Value = MaxIncrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
        ws.Cells(4, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = MaxDecreaseTicker
        ws.Cells(4, 16).Value = MaxDecrease
        ws.Cells(4, 16).NumberFormat = "0.00%"
        
        ws.Cells(5, 14).Value = "Greatest Total Volume"
        ws.Cells(5, 15).Value = MaxVolumeTicker
        ws.Cells(5, 16).Value = MaxVolume

    Next ws
    
    ' Display a message box when the macro is done running.
    
    MsgBox "Worksheets have been analyzed!"

End Sub
