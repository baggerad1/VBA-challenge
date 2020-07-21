# VBA-challenge

Sub Stock()
    
    
      Dim WS As Worksheet
        For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
        ' What is the last row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' need a header for the rows
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        'need something to hold values
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        'open price needs to be set
        Open_Price = Cells(2, Column + 2).Value
         ' Loop through all ticker symbol
        
        For i = 2 To LastRow
         ' This is challenging, need to look if we are on the ticker symbol
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ' Ticker name
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
                ' Close Price
                Close_Price = Cells(i, Column + 5).Value
                ' Yearly Change
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 9).Value = Yearly_Change
                ' I need to add Percent Change/ Getting to used to Python remember the Elif and Elseif difference
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Total Volue
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                ' Summary table row
                Row = Row + 1
                ' reset opening price
                Open_Price = Cells(i + 1, Column + 2)
                ' Volume Total reset
                Volume = 0
            'if ticker is same
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' What was the Last Row of Yearly Change per
        Yearly_Change = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' End here for today. Pick up tomorrow on Cell Colors(go to youtube and watch the Everyday VBA version
        For j = 2 To Yearly_Change
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        ' Find the ticker
        For Z = 2 To Yearly_Change
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & Yearly_Change)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & Yearly_Change)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & Yearly_Change)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
        
    Next WS
        
End Sub

