Dim Summary_Stocks As Double
Sub GetTicker()

  ' Define Variables

    Dim ws As Worksheet
    Dim Ticker As String
    Dim LastRow As Long
    Dim OpenValue As Double
    Dim LastValue As Double
    Dim YearlyChange As Double
    Dim PerChange As Double
    Dim Total_Stocks As Double
    Dim Increase As Double
    Dim Decrease As Double
    Dim Volume As Double
    Dim IncreaseTicker As String
    Dim DecreaseTicker As String
    Dim Volumeticker As String
    
    For Each ws In Worksheets
        ' Intialise the Variables
        Total_Stocks = 0
  
        ' Build the Summary Table
        Summary_Stocks = 2
    
        Dim firstRow As Long
        firstRow = 2
  
        'Define the Header of the Summary Table
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "YearlyChange"
        ws.Cells(1, 13).Value = "PerChange"
        ws.Cells(1, 14).Value = "Total_Stocks"
        
  
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' Loop through all Rows
        For i = 2 To LastRow + 1
     
            ' Check if we are still within the same Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                ' Get the Ticker name
                Ticker = ws.Cells(i, 1).Value
    
                ' Get the Total Stocks
                Total_Stocks = Total_Stocks + ws.Cells(i, 7).Value
          
                'Get the Last Value,Open Value and then Compute the Yearly Change and Percentage
                LastValue = ws.Cells(i, 6).Value
                OpenValue = ws.Cells(firstRow, 6).Value
                YearlyChange = LastValue - OpenValue
          
                If OpenValue <> 0 Then
                    PerChange = YearlyChange / OpenValue
                Else
                    PerChange = YearlyChange / 100
                End If
    
                ' Print the Ticker in the Summary Table
                ws.Range("K" & Summary_Stocks).Value = Ticker
    
                ' Print the Total Stocks to the Summary Table
                ws.Range("N" & Summary_Stocks).Value = Total_Stocks
          
                ' Print the YearlyChange to the Summary Table
                ws.Range("L" & Summary_Stocks).Value = YearlyChange
                
                ' Print the PerChange to the Summary Table
                ws.Range("M" & Summary_Stocks).Value = PerChange
          
                'Format the PerChange
                ws.Range("M" & Summary_Stocks).NumberFormat = "0.00%"
          
                'Colour Format the YearlyChange
                If YearlyChange < 0 Then
                    ws.Range("L" & Summary_Stocks).Interior.ColorIndex = 3
                Else
                    ws.Range("L" & Summary_Stocks).Interior.ColorIndex = 4
                End If
    
                ' Add one to the summary table row
                Summary_Stocks = Summary_Stocks + 1
          
                ' Reset the Brand Total
                Total_Stocks = 0
                
                firstRow = i + 1
          
            ' If the cell immediately following a row is the same Ticker...
            Else
    
                ' Add to the Total Stocks
                Total_Stocks = Total_Stocks + ws.Cells(i, 7).Value
            End If
        Next i

    
        ' Assign Last Row of Ticker Data Set
        LastRow = Summary_Stocks
      
        ' Define the Header of the Summary Table
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest %Increase"
        ws.Cells(3, 16).Value = "Greatest %Decrease"
        ws.Cells(4, 16).Value = "Greatest TotalVolume"
        
        Increase = ws.Cells(2, 13).Value
        IncreaseTicker = ws.Cells(2, 11).Value
        
        Decrease = ws.Cells(2, 13).Value
        DecreaseTicker = ws.Cells(2, 11).Value
        
        Volume = ws.Cells(2, 14).Value
        Volumeticker = ws.Cells(2, 11).Value
   
        ' Loop through all Rows
        For i = 2 To LastRow
            If (Increase < ws.Cells(i + 1, 13).Value) Then
                Increase = ws.Cells(i + 1, 13).Value
                IncreaseTicker = ws.Cells(i + 1, 11).Value
            End If
            
            If (Decrease > ws.Cells(i + 1, 13).Value) Then
                Decrease = ws.Cells(i + 1, 13).Value
                DecreaseTicker = ws.Cells(i + 1, 11).Value
            End If
            
            If (Volume < ws.Cells(i + 1, 14).Value) Then
                Volume = ws.Cells(i + 1, 14).Value
                Volumeticker = ws.Cells(i + 1, 11).Value
            End If
        Next i
    
        ' Print Values in Stock Stat Table
        ws.Cells(2, 17).Value = IncreaseTicker
        ws.Cells(2, 18).Value = Increase
        ws.Range("R" & 2).NumberFormat = "0.00%"
        ws.Cells(3, 17).Value = DecreaseTicker
        ws.Cells(3, 18).Value = Decrease
        ws.Range("R" & 3).NumberFormat = "0.00%"
        ws.Cells(4, 17).Value = Volumeticker
        ws.Cells(4, 18).Value = Volume

    Next

End Sub



