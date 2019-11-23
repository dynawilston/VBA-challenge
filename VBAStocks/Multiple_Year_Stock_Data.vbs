Attribute VB_Name = "Module2"
Sub Alphabetical_Testing()

For Each ws In Worksheets

Dim Total_Ticker_Volume As Double
Dim Ticker As String
Dim Yearly_Open As Double
Yearly_Open = 0
Dim Yearly_Close As Double
Yearly_Close = 0
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim SummaryRow As Long
Dim Open_Row As Long
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total_Volume As Double

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Value"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


Total_Ticker_Volume = 0
Open_Row = 2
SummaryRow = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
LastRow_ = ws.Cells(Rows.Count, 11).End(xlUp).Row


For i = 2 To LastRow
    
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
    ws.Range("L" & SummaryRow).Value = Total_Ticker_Volume
    
    
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    Ticker_Symbol = ws.Cells(i, 1).Value
    ws.Range("I" & SummaryRow).Value = Ticker_Symbol
    
    Yearly_Open = Yearly_Open + ws.Range("C" & Open_Row).Value
    Yearly_Close = Yearly_Close + ws.Range("F" & i).Value
    Yearly_Change = Yearly_Close - Yearly_Open
    ws.Range("J" & SummaryRow).Value = Yearly_Change
    
        If Yearly_Open = 0 Then
        Percent_Change = 0
        
        Else
        Percent_Change = Yearly_Change / Yearly_Open
        ws.Range("K" & SummaryRow).Value = Percent_Change
        ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
        End If
    
    If ws.Range("J" & SummaryRow).Value > 0 Then
            ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
         Else
            ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
         End If

SummaryRow = SummaryRow + 1
Open_Row = i + 1
Total_Ticker_Volume = 0
End If

Next i

    For i = 2 To LastRow_
    
     
    If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
    ws.Range("P2").Value = ws.Range("I" & i).Value
    ws.Range("Q2").Value = ws.Range("K" & i).Value
    
    
    End If
    
    If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
    ws.Range("Q3").Value = ws.Range("K" & i).Value
    ws.Range("P3").Value = ws.Range("I" & i).Value
    
    End If
    
    If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
    ws.Range("Q4").Value = ws.Range("L" & i).Value
    ws.Range("P4").Value = ws.Range("I" & i).Value
    End If
    
    
    
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "##0.0E+0"
        ws.Columns("I:Q").AutoFit
 
 Next i
 

Next ws

End Sub


