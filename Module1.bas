Attribute VB_Name = "Module1"
Sub YearStockData()

    For Each ws In Worksheets

    Dim i As Long
    Dim m As Long
    
    Dim Ticker_Name As String
    Dim Ticker_Title As String
    
    Dim QuarterlyChange As Double
    Dim Quarterly_Change_Tittle As String
    
    Dim PerccentChange As Double
    Dim Percent_Change_Tittle As String
        
    Dim Vol_Total As Double
        Vol_Total = 0
    Dim Vol_Title As String
    
    Dim GreatInc As String
    Dim GreatDec As String
    Dim GreatVol As String
    
    Dim MaxInc As Double
    Dim MaxDec As Double
    Dim MaxVol As Double
                 
    ' Starting Row for Table Creation
    Dim StartRow As Long
        StartRow = 2
          
    'Starting Row for calculation
    Dim j As Long
        j = 2
      
    ' Starting Row for Table Creation
    Dim Start As Long
        Start = 2
    
    WorksheetName = ws.Name
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Defining the Headers and Labels
    
    Ticker_Tittle = "Ticker"
    ws.Range("I1").Value = Ticker_Tittle
    
    Quarterly_Change_Tittle = "Quarterly Change"
    ws.Range("J1").Value = Quarterly_Change_Tittle
    
    Percent_Change_Tittle = "Percentage Change"
    ws.Range("K1").Value = Percent_Change_Tittle
    
    Vol_Tittle = "Volume_Total"
    ws.Range("L1").Value = Vol_Tittle
    
    GreatInc = "Greatest % Increase"
    ws.Cells(2, 15).Value = GreatInc
    
    GreatDec = "Greatest % Decrease"
    ws.Cells(3, 15).Value = GreatDec
    
    GreatVol = "Greatest Total Volume"
    ws.Cells(4, 15).Value = GreatVol
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
               
    
                   
    'Loop 1 initial Calculations and List
        For i = 2 To LastRow
     
     'Placing the Ticker List
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker_Name = ws.Cells(i, 1).Value
              
    'Quarterly Change Calculation
            QuarterlyChange = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                If ws.Range("J" & StartRow).Value < 0 Then
                ws.Range("J" & StartRow).Interior.ColorIndex = 3
                ElseIf ws.Range("J" & StartRow).Value > 0 Then
                ws.Range("J" & StartRow).Interior.ColorIndex = 4
                Else
                ws.Range("J" & StartRow).Interior.ColorIndex = xlNone
                End If
                     
     'Percent Change Calculation
                If ws.Cells(j, 3).Value <> 0 Then
                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / (ws.Cells(j, 3).Value))
                
                    With ws.Range("K" & StartRow)
                    .NumberFormat = "0.00%"
                    .Value = PercentChange
                    End With
                
                Else
                ws.Cells(StartRow, 11).Value = 0
                End If
            
     'Sum of Volume Calculation per quarter
           
            Vol_Total = Vol_Total + ws.Cells(i, 7).Value
                    
            ws.Range("I" & StartRow).Value = Ticker_Name
            ws.Range("J" & StartRow).Value = QuarterlyChange
            ws.Range("L" & StartRow).Value = Vol_Total
            
            StartRow = StartRow + 1
            
            j = i + 1
            
            Vol_Total = 0
        
            Else
                            
            Vol_Total = Vol_Total + ws.Cells(i, 7).Value
            
            End If
        
        Next i
    
    'Greatest Increase, Decrease and Volume Calculation
    
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    
        MaxInc = ws.Cells(2, 11).Value
        MaxDec = ws.Cells(2, 11).Value
        MaxVol = ws.Cells(2, 12).Value
      
        For m = 2 To LastRowI
    
            If ws.Cells(m + 1, 11).Value > MaxInc Then
                MaxInc = ws.Cells(m + 1, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(m + 1, 9).Value
                    With ws.Cells(2, 17)
                    .NumberFormat = "0.00%"
                    .Value = MaxInc
                    End With
            Else
                MaxInc = MaxInc
            End If
        
            If ws.Cells(m + 1, 12).Value > MaxVol Then
                MaxVol = ws.Cells(m + 1, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(m + 1, 9).Value
                    With ws.Cells(4, 17)
                    .NumberFormat = "#.#0E+0"
                    .Value = MaxVol
                    End With
            Else
                MaxVol = MaxVol
            End If
        
            If ws.Cells(m + 1, 11).Value < MaxDec Then
                MaxDec = ws.Cells(m + 1, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(m + 1, 9).Value
                    With ws.Cells(3, 17)
                    .NumberFormat = "0.00%"
                    .Value = MaxDec
                    End With
            Else
                MaxDec = MaxDec
            End If
        
           
        Next m
    
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
    
    Next ws
    
            
End Sub
