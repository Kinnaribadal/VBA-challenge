Attribute VB_Name = "Module1"
Sub Volume()
' Create a variable to hold ticker symbol
  Dim Ticker As String

  ' Create a variable to hold opening, closing, stock volume
  Dim Opening As Double
  Dim Closing As Double
  Dim Volume As Double
  Dim PercentChange As Double
  Dim Change As Double
  

    Volume = 0
    j = 0
    StartValue = 2
    

    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly change"
    Range("L1").Value = "Percent change"
    Range("M1").Value = "Total Volume"
    
  RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  
  For i = 2 To RowCount
    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        Volume = Volume + Cells(i, 7).Value
        
        'to find non zero value
        If Cells(StartValue, 3) = 0 Then
            For FindValue = StartValue To i
                If Cells(FindValue, 3).Value <> 0 Then
                    StartValue = FindValue
                    Exit For
                End If
            Next FindValue
        End If
        
        Change = Cells(i, 6) - Cells(StartValue, 3)
        PercentChange = (Change / Cells(StartValue, 3)) * 100
        
        
        Range("J" & 2 + j).Value = Cells(i, 1)
        Range("K" & 2 + j).Value = Change
        Range("L" & 2 + j).Value = PercentChange
        Range("M" & 2 + j).Value = Volume
        If Change < 0 Then
            Range("K" & 2 + j).Interior.ColorIndex = 3
        Else
            Range("K" & 2 + j).Interior.ColorIndex = 4
        End If
            
            
        Volume = 0
        j = j + 1
        Change = 0
        
    Else
        Volume = Volume + Cells(i, 7).Value
    End If
    
    Next i
        


End Sub
