Attribute VB_Name = "Stock_exchange"

Sub Stockexchange()

Dim WS As Worksheet
For Each WS In Worksheets
    WS.Activate

Dim i As Long
Dim ticker As String
Dim volume As Double
    volume = 0
Dim summary_ticker_row As Integer
    summary_ticker_row = 2

Dim opening_price As Double
        opening_price = Cells(2, 3).Value
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
        
        Cells(1, 9).Value = "Summary Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

   Dim lastrow As Long
   
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow

            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
              ticker = Cells(i, 1).Value
              volume = volume + Cells(i, 7).Value
              Range("I" & summary_ticker_row).Value = ticker
              Range("L" & summary_ticker_row).Value = volume
              
                closing_price = Cells(i, 6).Value
                yearly_change = (closing_price - opening_price)
              
              Range("J" & summary_ticker_row).Value = yearly_change

                If opening_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = (yearly_change / opening_price) - 1
                End If

                 Range("K" & summary_ticker_row).Value = percent_change
                summary_ticker_row = summary_ticker_row + 1
              
              volume = 0
              opening_price = Cells(i + 1, 3)
            
            Else
         
              volume = volume + Cells(i, 7).Value

            
            End If
        
        Next i

Range("I1:L1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"

    lastrow_summary_ticker = Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To lastrow_summary_ticker
            
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.Color = vbGreen
                
            Else
                Cells(i, 10).Interior.Color = vbRed
                
            End If

        Next i

Columns("L:L").Select
Selection.NumberFormat = "0"
Range("A1").Select

Next

End Sub


