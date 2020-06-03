Attribute VB_Name = "Module1"
Sub VBA_Stocks()
    
    Sheets(1).Select
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Volume_Ticker As String
    Dim Greatest_Increase_Value As Double
    Dim Greatest_Decrease_Value As Double
    Dim Greatest_Volume_Value As Double
        
    Greatest_Increase_Value = 0
    Greatest_Decrease_Value = 0
    Greatest_Volume_Value = 0
                            
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim Ticker As String
        
        Dim Opening_Value As Double
        Opening_Value = Cells(2, 3).Value
        
        Dim Closing_Value As Double
        
        Dim Stock_Volume As Double
        Stock_Volume = 0

        Dim Results_Row As Integer
        Results_Row = 2
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

            For i = 2 To LastRow
                
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                    Closing_Value = Cells(i, 6).Value
                    
                    Range("J" & Results_Row).Value = Closing_Value - Opening_Value
                        
                        If Opening_Value = 0 Then
                        
                        Range("K" & Results_Row).Value = 0
                    
                        ElseIf ((Closing_Value - Opening_Value) / Opening_Value) > 0 Then
                        
                        Range("K" & Results_Row).Value = (Closing_Value - Opening_Value) / Opening_Value

                        Range("K" & Results_Row).Interior.ColorIndex = 4
                        
                        Else
                        
                        Range("K" & Results_Row).Value = (Closing_Value - Opening_Value) / Opening_Value
                        
                        Range("K" & Results_Row).Interior.ColorIndex = 3
                    
                        End If
                       
                        Opening_Value = Cells(i + 1, 3).Value
                    
                    Ticker = Cells(i, 1).Value
        
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    
                    Range("I" & Results_Row).Value = Ticker

                    Range("L" & Results_Row).Value = Stock_Volume

                    Results_Row = Results_Row + 1
      
                    Stock_Volume = 0
                                        
                 Else
                    
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    
                End If
                
            Next i
        
        Range("K2:K" & LastRow).NumberFormat = "0.00%"
        
        Dim NewLastRow As Long
        NewLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
            For j = 2 To NewLastRow
            
                If Cells(j, 11).Value > Greatest_Increase_Value Then
            
                Greatest_Increase_Value = Cells(j, 11).Value
            
                Greatest_Increase_Ticker = Cells(j, 9).Value
            
                ElseIf Cells(j, 11).Value < Greatest_Decrease_Value Then
            
                Greatest_Decrease_Value = Cells(j, 11).Value
            
                Greatest_Decrease_Ticker = Cells(j, 9).Value
                
                End If
                        
                If Cells(j, 12).Value > Greatest_Volume_Value Then
                        
                Greatest_Volume_Value = Cells(j, 12).Value
                
                Greatest_Volume_Ticker = Cells(j, 9).Value
                
                End If
                                                
            Next j
    
    Next ws
        Sheets(1).Select
            Range("Q2:Q3").NumberFormat = "0.00%"
            Range("P2").Value = Greatest_Increase_Ticker
            Range("P3").Value = Greatest_Decrease_Ticker
            Range("P4").Value = Greatest_Volume_Ticker
            Range("Q2").Value = Greatest_Increase_Value
            Range("Q3").Value = Greatest_Decrease_Value
            Range("Q4").Value = Greatest_Volume_Value

End Sub
