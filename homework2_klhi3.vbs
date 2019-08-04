Sub calculateV()

  'worksheet
  Dim ws_count As Integer
  Dim wi As Integer
  Dim sheet1 As Worksheet
  
  ws_count = ActiveWorkbook.Worksheets.Count
  
  For wi = 1 To ws_count
      Set sheet1 = ActiveWorkbook.Worksheets(wi)

  
  'working on sheet
  
  Dim i As Long
  Dim xsum As Double
  Dim x As String
  Dim first_value As Double
  Dim end_value As Double
  
  
' Header
  sheet1.Range("I1").Value = "Ticker"
  sheet1.Range("J1").Value = "Yearly Change"
  sheet1.Range("K1").Value = "Percent Change"
  sheet1.Range("L1").Value = "Total Stock Volume"
  rangerow = 2
    
' Last row
  lastrow = sheet1.Cells(Rows.Count, 1).End(xlUp).Row
  lastrow1 = lastrow - 1
  
' First Value
  x = sheet1.Cells(2, 1).Value
  xsum = sheet1.Cells(2, 7).Value
  first_value = sheet1.Cells(2, 3).Value
  end_value = sheet1.Cells(2, 6).Value
 

'Check rows except last row
  For i = 3 To lastrow1
  
   If (sheet1.Cells(i, 3).Value <> 0) Then
     If (sheet1.Cells(i, 1).Value = x) Then
        xsum = xsum + sheet1.Cells(i, 7).Value
        end_value = sheet1.Cells(i, 6).Value
     Else
        sheet1.Range("I" & rangerow).Value = x
        sheet1.Range("L" & rangerow).Value = xsum
        sheet1.Range("J" & rangerow).Value = end_value - first_value
        
        If first_value <> 0 Then
           sheet1.Range("K" & rangerow).Value = (end_value - first_value) / first_value
        Else
           sheet1.Range("K" & rangerow).Value = end_value
        End If
        
        'assign new values
        x = sheet1.Cells(i, 1).Value
        xsum = sheet1.Cells(i, 7).Value
        rangerow = rangerow + 1
        
        first_value = sheet1.Cells(i, 3).Value
        end_value = sheet1.Cells(i, 6).Value
        
     End If
    End If  'open value <>0
  Next i
  
'last row
  If (sheet1.Cells(lastrow, 1).Value = x) Then
        xsum = xsum + sheet1.Cells(lastrow, 7).Value
        
        sheet1.Range("I" & rangerow).Value = x
        sheet1.Range("L" & rangerow).Value = xsum
        sheet1.Range("J" & rangerow).Value = end_value - first_value
        If first_value <> 0 Then
           sheet1.Range("K" & rangerow).Value = (end_value - first_value) / first_value
        Else
           sheet1.Range("K" & rangerow).Value = end_value
        End If
        
  Else
    sheet1.Range("I" & rangerow).Value = x
    sheet1.Range("L" & rangerow).Value = xsum
    
    end_value = sheet1.Cells(rangerow, 6).Value
    sheet1.Range("J" & rangerow).Value = end_value - first_value
        
    If first_value <> 0 Then
        sheet1.Range("K" & rangerow).Value = (end_value - first_value) / first_value
    Else
        sheet1.Range("K" & rangerow).Value = end_value
    End If
  End If
  
  
  Dim min_ticker, max_ticker, vol_ticker As String
  Dim min_value, max_value, vol_max As Double
  
  
  'min, max, total
  min_ticker = sheet1.Range("I2").Value
  min_value = sheet1.Range("K2").Value
  max_ticker = sheet1.Range("I2").Value
  max_value = sheet1.Range("K2").Value
  vol_max = sheet1.Range("L2").Value
  vol_ticker = sheet1.Range("I2").Value
  
  
  For i = 3 To lastrow
    If sheet1.Cells(i, 11).Value > max_value Then
        max_value = sheet1.Cells(i, 11).Value
        max_ticker = sheet1.Cells(i, 9).Value
    End If
    
    If sheet1.Cells(i, 11).Value < min_value Then
        min_value = sheet1.Cells(i, 11).Value
        min_ticker = sheet1.Cells(i, 9).Value
    End If
    
    If sheet1.Cells(i, 12).Value > vol_max Then
        vol_max = sheet1.Cells(i, 12).Value
        vol_ticker = sheet1.Cells(i, 9).Value
    End If
  Next i
  
  sheet1.Range("O1").Value = sheet1.Name
  
  sheet1.Range("P1").Value = "Ticker"
  sheet1.Range("Q1").Value = "Value"
  sheet1.Range("O2").Value = "Greatest % Increase"
  sheet1.Range("P2").Value = max_ticker
  sheet1.Range("Q2").Value = max_value
  
  sheet1.Range("O3").Value = "Greatest % Decrease"
  sheet1.Range("P3").Value = min_ticker
  sheet1.Range("Q3").Value = min_value
  
  sheet1.Range("O4").Value = "Greatest Total Volume"
  sheet1.Range("P4").Value = vol_ticker
  sheet1.Range("Q4").Value = vol_max
    
    
  Next wi   'sheet
  
End Sub

