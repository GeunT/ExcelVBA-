Sub Multi_Year_testing()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

    Dim Name_Ticker As String
    Dim Total_Volume As Double
    Dim Yearly_Change As Variant
    Dim Percent_Change As Variant
    Dim Lastrow As Long
    Dim j As Long
    Dim g As Long

    Dim OpenValue As Variant
    Dim CloseValue As Variant
    
    Dim annualdifference As Double
 
    Total_Volume = 0
    
Dim Summary_Test As Double
Summary_Test = 2

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Cells(1, 18).Value = "Open Price"
    Cells(1, 19).Value = "Close Price"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % increase"
    Cells(3, 14).Value = "Greatest % decrease"
    Cells(4, 14).Value = "Largest Volume"
    
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    j = 2
    annualdifference = Cells(2, 3).Value
    
    For I = 2 To Lastrow
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            Name_Ticker = Cells(I, 1).Value
            Total_Volume = Total_Volume + Cells(I, 7).Value
            Yearly_Change = Cells(I, 6).Value - annualdifference
                
                If annualdifference = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = Format((Cells(I, 6).Value - annualdifference) / annualdifference, "0.00%")
                End If
                
             j = j + 1
             
             annualdifference = Cells(I + 1, 3).Value
                
                If OpenValue <> 0 And CloseValue = 0 Then
                CloseValue = Cells(I, 6).Value
                  End If
            
            Range("R" & Summary_Test).Value = OpenValue
            Range("S" & Summary_Test).Value = CloseValue
            
            Range("I" & Summary_Test).Value = Name_Ticker
            Range("J" & Summary_Test).Value = Yearly_Change
            Range("K" & Summary_Test).Value = Percent_Change
            Range("L" & Summary_Test).Value = Total_Volume
        
            Summary_Test = Summary_Test + 1
            Total_Volume = 0
         
    Else
    
            Total_Volume = Total_Volume + Cells(I, 7).Value
             
             If OpenValue = 0 Then
                OpenValue = Cells(I, 3).Value
                End If
            
           If CloseValue > 0 And OpenValue > 0 Then
                OpenValue = 0
                CloseValue = 0
                End If
            End If
               
        If Cells(I, 10).Value > 0 Then
        Cells(I, 10).Interior.ColorIndex = 3
        Else: Cells(I, 10).Interior.ColorIndex = 4
        End If
        
    Next I

    TLastrow = Cells(Rows.Count, 9).End(xlUp).Row
        For g = 2 To TLastrow
                If Cells(g, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & TLastrow)) Then
                    Cells(2, 15).Value = Cells(g, 9).Value
                    Cells(2, 16).Value = Cells(g, 11).Value
                    Cells(2, 16).NumberFormat = "0.00%"
                ElseIf Cells(g, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & TLastrow)) Then
                    Cells(3, 15).Value = Cells(g, 9).Value
                    Cells(3, 16).Value = Cells(g, 11).Value
                    Cells(3, 16).NumberFormat = "0.00%"
                ElseIf Cells(g, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & TLastrow)) Then
                    Cells(4, 15).Value = Cells(g, 9).Value
                    Cells(4, 16).Value = Cells(g, 12).Value
                End If
            Next g
 
Next ws

    MsgBox ("complete")
End Sub




