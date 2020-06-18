Attribute VB_Name = "Module1"

Sub stock_doc()

    For Each ws In ActiveWorkbook.Worksheets


    Dim lastRow As Long
    Dim ticker As String
    Dim Volume As Variant
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim yearChange As Single
    Dim percentChange As Variant
    Dim rowCounter As Variant
    Dim resultsCounter As Variant
 
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    resultsCounter = 2

     lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Volume = 0

    yearOpen = ws.Cells(2, 3).Value
    
    For rowCounter = 2 To lastRow
        
        Volume = Volume + ws.Cells(rowCounter, 7).Value
        
        If (ws.Cells(rowCounter - 1, 1).Value = ws.Cells(rowCounter, 1).Value And ws.Cells(rowCounter + 1, 1).Value <> ws.Cells(rowCounter, 1).Value) Then
           
           yearClose = ws.Cells(rowCounter, 6).Value

           yearChange = yearClose - yearOpen

           If yearOpen = 0 And yearClose <> 0 Then

                percentChange = yearClose / yearClose

           ElseIf yearOpen = 0 And yearClose = 0 Then

                percentChange = 0
            
           Else

                percentChange = (yearClose - yearOpen) / yearOpen

           End If

           yearOpen = ws.Cells(rowCounter + 1, 3).Value
           
            ws.Cells(resultsCounter, 10).Value = yearChange

            If ws.Cells(resultsCounter, 10).Value < 0 Then

                ws.Cells(resultsCounter, 10).Interior.ColorIndex = 3
            Else

                ws.Cells(resultsCounter, 10).Interior.ColorIndex = 4
           
        End If
          
           ws.Cells(resultsCounter, 11).Value = percentChange
           
           ws.Cells(resultsCounter, 11).NumberFormat = "0.00%"
          
        End If

        If (ws.Cells(rowCounter + 1, 1).Value <> ws.Cells(rowCounter, 1).Value) Then
            
            ws.Cells(resultsCounter, 9).Value = ws.Cells(rowCounter, 1).Value
        
            ws.Cells(resultsCounter, 12).Value = Volume

            Volume = 0

           resultsCounter = resultsCounter + 1

        End If
       
    Next rowCounter

    ws.Cells(2, 17).Value = 0
    ws.Cells(3, 17).Value = 0
    ws.Cells(4, 17).Value = 0

    For resultsCounter = 2 To lastRow

        If ws.Cells(resultsCounter, 11).Value > ws.Cells(2, 17).Value Then

            ws.Cells(2, 17).Value = ws.Cells(resultsCounter, 11).Value

            ws.Cells(2, 16).Value = ws.Cells(resultsCounter, 9).Value
            
        End If

        If ws.Cells(resultsCounter, 11).Value < ws.Cells(3, 17).Value Then

            ws.Cells(3, 17).Value = ws.Cells(resultsCounter, 11).Value

            ws.Cells(3, 16).Value = ws.Cells(resultsCounter, 9).Value
        
        End If

         If ws.Cells(resultsCounter, 12).Value > ws.Cells(4, 17).Value Then

            ws.Cells(4, 17).Value = ws.Cells(resultsCounter, 12).Value

            ws.Cells(4, 16).Value = ws.Cells(resultsCounter, 9).Value
            
        End If
       
        Next resultsCounter
        
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"


    Next ws
 
End Sub
