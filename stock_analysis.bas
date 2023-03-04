Attribute VB_Name = "assignment2"
Sub stockAnalysis_Fine()
 'create workbook sheets variable
    Dim wsCount As Integer
    
    'determine number of worksheets
    wsCount = ActiveWorkbook.Worksheets.Count
    
    'Worksheet For loop
    For s = 1 To wsCount
  
         'set column header
           ActiveWorkbook.Worksheets(s).Range("K1").Value = "Ticker"
           ActiveWorkbook.Worksheets(s).Range("L1").Value = "Yearly Change"
           ActiveWorkbook.Worksheets(s).Range("M1").Value = "Percentage Change"
           ActiveWorkbook.Worksheets(s).Range("N1").Value = "Total Volume"
           
           'determine last row
           Dim lastRow As Long
           lastRow = ActiveWorkbook.Worksheets(s).Range("A1").End(xlDown).row
           
           'set variables
           Dim yearlyChange As Variant
           Dim firstOfYear As Variant
           Dim lastOfYear As Variant
           Dim percentChange As Variant
           Dim totalVolum As Long
           Dim valuePlace As Variant
           
           'start unique value placement at row 2
           valuePlace = 2
           
           
           For i = 2 To lastRow
               If ActiveWorkbook.Worksheets(s).Cells(i + 1, 1).Value = ActiveWorkbook.Worksheets(s).Cells(i, 1).Value And Right(ActiveWorkbook.Worksheets(s).Cells(i, 2), 4) = "0102" Then
                   firstOfYear = ActiveWorkbook.Worksheets(s).Cells(i, 3).Value
                   totalVolume = ActiveWorkbook.Worksheets(s).Cells(i, 7).Value
                   
                   ElseIf ActiveWorkbook.Worksheets(s).Cells(i + 1, 1).Value = ActiveWorkbook.Worksheets(s).Cells(i, 1).Value Then
                       totalVolume = totalVolume + Cells(i, 7).Value
                       
                   Else
                       lastOfYear = ActiveWorkbook.Worksheets(s).Cells(i, 3).Value
                       totalVolume = totalVolume + ActiveWorkbook.Worksheets(s).Cells(i, 7).Value
                       yearlyChange = lastOfYear - firstOfYear
                       percentChange = yearlyChange / firstOfYear
                       
                       'set ticker
                       ActiveWorkbook.Worksheets(s).Range("K" & valuePlace).Value = ActiveWorkbook.Worksheets(s).Cells(i, 1).Value
                       'set yearly change
                       ActiveWorkbook.Worksheets(s).Cells(valuePlace, 12).Value = yearlyChange
                       'set percentage change
                       ActiveWorkbook.Worksheets(s).Cells(valuePlace, 13).Value = FormatPercent(percentChange)
                       'set total volume
                       ActiveWorkbook.Worksheets(s).Cells(valuePlace, 14).Value = totalVolume
                       'increase row placement for ticker
                       valuePlace = valuePlace + 1
               End If
           Next i
           
        'set format of yearly change and percentage change
        
        'determine last row of unique tickers
        Dim lastTick As Long
        lastTick = ActiveWorkbook.Worksheets(s).Range("K2").End(xlDown).row
        
           For i = 2 To lastTick
               If ActiveWorkbook.Worksheets(s).Cells(i, 12).Value > 0 Then
                   ActiveWorkbook.Worksheets(s).Cells(i, 12).Interior.ColorIndex = 4
                   
               Else
                   ActiveWorkbook.Worksheets(s).Cells(i, 12).Interior.ColorIndex = 3
                   
               End If
           Next i
           
                   
           'table creation to determin greatest stocks
           ActiveWorkbook.Worksheets(s).Range("R1").Value = "Ticker"
           ActiveWorkbook.Worksheets(s).Range("S1").Value = "Value"
           ActiveWorkbook.Worksheets(s).Range("Q2").Value = "Greatest % Increase"
           ActiveWorkbook.Worksheets(s).Range("Q3").Value = "Greatest % Decrease"
           ActiveWorkbook.Worksheets(s).Range("Q4").Value = "Greatest Total Volume"
           
           
           
           'Determine Greatest % Increast
           Dim percentIncrease As Variant
           Dim tickerIncrease As String
           
           percentIncrease = 0
           
           For i = 2 To lastTick
               If ActiveWorkbook.Worksheets(s).Cells(i, 13).Value > percentIncrease Then
                   percentIncrease = ActiveWorkbook.Worksheets(s).Cells(i, 13).Value
                   tickerIncrease = ActiveWorkbook.Worksheets(s).Cells(i, 11).Value
               End If
           Next i
           ActiveWorkbook.Worksheets(s).Range("R2").Value = tickerIncrease
           ActiveWorkbook.Worksheets(s).Range("S2").Value = FormatPercent(percentIncrease)
               
           'Determine Greatest % Decrease
           Dim percentDecrease As Variant
           Dim tickerDecrease As String
           
           percentDecrease = 0
           
           For i = 2 To lastTick
               If ActiveWorkbook.Worksheets(s).Cells(i, 13).Value < percentDecrease Then
                   percentDecrease = ActiveWorkbook.Worksheets(s).Cells(i, 13).Value
                   tickerDecrease = ActiveWorkbook.Worksheets(s).Cells(i, 11).Value
               End If
           Next i
           ActiveWorkbook.Worksheets(s).Range("R3").Value = tickerDecrease
           ActiveWorkbook.Worksheets(s).Range("S3").Value = FormatPercent(percentDecrease)
           
           'determine greatest volume
           Dim greatVolume As Variant
           Dim tickerVolume As String
           
           greatVolume = 0
           
           For i = 2 To lastTick
               If ActiveWorkbook.Worksheets(s).Cells(i, 14).Value > greatVolume Then
                   greatVolume = ActiveWorkbook.Worksheets(s).Cells(i, 14).Value
                   tickerVolume = ActiveWorkbook.Worksheets(s).Cells(i, 11).Value
               End If
           Next i
           ActiveWorkbook.Worksheets(s).Range("R4").Value = tickerVolume
           ActiveWorkbook.Worksheets(s).Range("S4").Value = greatVolume
        Worksheets(s).Columns("K:S").AutoFit
        MsgBox ("Worksheet " & s & " is done!")
    Next s
    MsgBox ("All Done")
End Sub

