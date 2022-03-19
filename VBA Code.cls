VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBAHW():

    For Each ws In Worksheets
    
        Dim WN As String
        Dim x As Long
        Dim y As Long
        Dim Count As Long
        Dim Row1 As Long
        Dim Row2 As Long
        Dim PercentChange As Double
        Dim GreatIncrease As Double
        Dim GreatDecrease As Double
        Dim GreatVolume As Double
        WN = ws.Name
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        Count = 2
        y = 2
        
        Row1 = ws.Cells(Rows.Count, 1).End(xlUp).Row

            For x = 2 To Row1

                If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then

                ws.Cells(Count, 9).Value = ws.Cells(x, 1).Value
                
                ws.Cells(Count, 10).Value = ws.Cells(x, 6).Value - ws.Cells(y, 3).Value
                
                    If ws.Cells(Count, 10).Value < 0 Then

                    ws.Cells(Count, 10).Interior.ColorIndex = 3
                
                    Else

                    ws.Cells(Count, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    If ws.Cells(y, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(x, 6).Value - ws.Cells(y, 3).Value) / ws.Cells(y, 3).Value)
                    
                    ws.Cells(Count, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(Count, 11).Value = Format(0, "Percent")
                    
                    End If
 
                ws.Cells(Count, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(y, 7), ws.Cells(x, 7)))

                Count = Count + 1

                y = x + 1
                
                End If
            
            Next x
            

        Row2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

            For x = 2 To Row2
            

                If ws.Cells(x, 12).Value > GreatVolume Then
                GreatVolume = ws.Cells(x, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(x, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
                
                End If

                If ws.Cells(x, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(x, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(x, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                
                End If

                If ws.Cells(x, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(x, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(x, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If

            ws.Cells(2, 17).Value = Format(GreatIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "Scientific")
            
            Next x
            
    Next ws
        
End Sub


