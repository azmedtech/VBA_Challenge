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

   
'assign headers to columns
Cells(1, "I").Value = "Ticker    "
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"
Cells(1, "P").Value = "Ticker    "
Cells(1, "Q").Value = "Value     "
Cells(2, "O").Value = "Greatest % Increase"
Cells(3, "O").Value = "Greatest % Decrease"
Cells(4, "O").Value = "Greatest Total Volume"
    
'fixed start definitions
OutputStartRow = 2
FirstRow = True
TSV = 0

    'begin for loop
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
                              
        'look for first row by ticker name, capture that opening value, then reset to not first row for next iteration
        If FirstRow = True Then
            OpenPrice = Cells(i, "C")
            FirstRow = False
        
        End If
                            
        'identify where the ticker name changes and write the values for Annual Change, Percent Change, and Total Stock Volume (TSV) into the appropriate locations
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, "A").Value
            Cells(OutputStartRow, "I").Value = Ticker
            ClosePrice = Cells(i, "F").Value
            AnnualChange = (ClosePrice - OpenPrice)
            Cells(OutputStartRow, "J").Value = AnnualChange
                              
            'accomodate for situations with division by zero
            If OpenPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = AnnualChange / OpenPrice
            End If
            
            'apply conditional formatting to cells in Yearly Change column
            If Cells(OutputStartRow, "J").Value > 0 Then
                Cells(OutputStartRow, "J").Interior.ColorIndex = 4
                Cells(OutputStartRow, "K").Interior.ColorIndex = 4
            Else
                Cells(OutputStartRow, "J").Interior.ColorIndex = 3
                Cells(OutputStartRow, "K").Interior.ColorIndex = 3
            End If
            
            Cells(OutputStartRow, "K").Value = PercentChange
            
            TSV = Cells(i, "G").Value + TSV
            Cells(OutputStartRow, "L").Value = TSV
                                           
            FirstRow = True
            OutputStartRow = OutputStartRow + 1
            
            TSV = 0
            
        Else
            TSV = Cells(i, "G").Value + TSV
        
        End If
                                                                                              
    Next i

'format columns
Range("I1").EntireColumn.AutoFit
Range("J1").EntireColumn.AutoFit
Range("K1").EntireColumn.NumberFormat = "0.00%"
Range("K1").EntireColumn.AutoFit
Range("L1").EntireColumn.AutoFit
Range("O1").EntireColumn.AutoFit


GreatestTable = Cells(Rows.Count, "I").End(xlUp).Row

    
    For i = 2 To GreatestTable
    
        If Cells(i, "K").Value = Application.WorksheetFunction.Max(Range("K2:K" & GreatestTable)) Then
                    Cells(2, "P").Value = Cells(i, "I").Value
                    Cells(2, "Q").Value = Cells(i, "K").Value
                    Cells(2, "Q").NumberFormat = "0.00%"
        
        ElseIf Cells(i, "K").Value = Application.WorksheetFunction.Min(Range("K2:K" & GreatestTable)) Then
            Cells(3, "P").Value = Cells(i, "I").Value
            Cells(3, "Q").Value = Cells(i, "K").Value
            Cells(3, "Q").NumberFormat = "0.00%"
        
        ElseIf Cells(i, "L").Value = Application.WorksheetFunction.Max(Range("L2:L" & GreatestTable)) Then
            Cells(4, "P").Value = Cells(i, "I").Value
            Cells(4, "Q").Value = Cells(i, "K").Value
        
        End If
        
    Next i
    
Range("P1").EntireColumn.AutoFit
Range("Q1:Q2").NumberFormat = "0.00%"
Range("Q1").EntireColumn.AutoFit
                    
End Sub

