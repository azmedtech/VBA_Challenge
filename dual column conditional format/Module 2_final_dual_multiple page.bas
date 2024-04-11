Attribute VB_Name = "Module1"
Sub financial_analysis()

'declare variables
Dim OutputStartRow As Integer
Dim FirstRow As Boolean
Dim i As Long
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim Ticker As String
Dim AnnualChange As Double
Dim PercentChange As Double
Dim TSV As LongLong
Dim GreatestTable As Integer

'begin worksheet loop
    For Each ws In Worksheets
           
        'assign headers to columns
        ws.Cells(1, "I").Value = "Ticker    "
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        ws.Cells(1, "P").Value = "Ticker    "
        ws.Cells(1, "Q").Value = "Value     "
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
            
        'fixed start definitions
        OutputStartRow = 2
        FirstRow = True
        TSV = 0
        
            'begin for loop
            For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
                                      
                'look for first row by ticker name, capture that opening value, then reset to not first row for next iteration
                If FirstRow = True Then
                    OpenPrice = Cells(i, "C")
                    FirstRow = False
                
                End If
                                    
                'identify where the ticker name changes and write the values for Annual Change, Percent Change, and Total Stock Volume (TSV) into the appropriate locations
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    Ticker = ws.Cells(i, "A").Value
                    ws.Cells(OutputStartRow, "I").Value = Ticker
                    ClosePrice = ws.Cells(i, "F").Value
                    AnnualChange = (ClosePrice - OpenPrice)
                    ws.Cells(OutputStartRow, "J").Value = AnnualChange
                                      
                    'accomodate for situations with division by zero
                    If OpenPrice = 0 Then
                        PercentChange = 0
                    Else
                        PercentChange = AnnualChange / OpenPrice
                    End If
                    
                    'apply conditional formatting to cells in Yearly Change column
                    If ws.Cells(OutputStartRow, "J").Value > 0 Then
                        ws.Cells(OutputStartRow, "J").Interior.ColorIndex = 4
                        ws.Cells(OutputStartRow, "K").Interior.ColorIndex = 4
                    Else
                        ws.Cells(OutputStartRow, "J").Interior.ColorIndex = 3
                        ws.Cells(OutputStartRow, "K").Interior.ColorIndex = 3
                    End If
                    
                    ws.Cells(OutputStartRow, "K").Value = PercentChange
                    
                    TSV = ws.Cells(i, "G").Value + TSV
                    ws.Cells(OutputStartRow, "L").Value = TSV
                                                   
                    FirstRow = True
                    OutputStartRow = OutputStartRow + 1
                    
                    TSV = 0
                    
                Else
                    TSV = ws.Cells(i, "G").Value + TSV
                
                End If
                                                                                                      
            Next i
        
        'format columns
        ws.Range("I1").EntireColumn.AutoFit
        ws.Range("J1").EntireColumn.AutoFit
        ws.Range("K1").EntireColumn.NumberFormat = "0.00%"
        ws.Range("K1").EntireColumn.AutoFit
        ws.Range("L1").EntireColumn.AutoFit
        ws.Range("O1").EntireColumn.AutoFit
        
        
        GreatestTable = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
            
            For i = 2 To GreatestTable
            
                If ws.Cells(i, "K").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & GreatestTable)) Then
                    ws.Cells(2, "P").Value = ws.Cells(i, "I").Value
                    ws.Cells(2, "Q").Value = ws.Cells(i, "K").Value
                    ws.Cells(2, "Q").NumberFormat = "0.00%"
        
                ElseIf ws.Cells(i, "K").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & GreatestTable)) Then
                    ws.Cells(3, "P").Value = ws.Cells(i, "I").Value
                    ws.Cells(3, "Q").Value = ws.Cells(i, "K").Value
                    ws.Cells(3, "Q").NumberFormat = "0.00%"
                
                ElseIf ws.Cells(i, "L").Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & GreatestTable)) Then
                    ws.Cells(4, "P").Value = ws.Cells(i, "I").Value
                    ws.Cells(4, "Q").Value = ws.Cells(i, "K").Value
                
                End If
                
            Next i
            
        Range("P1").EntireColumn.AutoFit
        Range("Q1:Q2").NumberFormat = "0.00%"
        Range("Q1").EntireColumn.AutoFit
                    
    Next ws
    
End Sub

