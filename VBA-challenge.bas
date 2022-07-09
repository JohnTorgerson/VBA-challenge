Attribute VB_Name = "Module1"
Sub VBA_Challenge()
    
    'define stock ticker as reference point
    Dim Ticker As String
    Dim LastRow As Long
    Dim i As Long
    Dim Ti As Long
        Ti = 2
        
    'create variables for open and close
    Dim openVal As Double
    Dim closeVal As Double
    Dim Volume As LongLong
      
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        openVal = Cells(2, 3).Value
        
    'Format width of gap cell
        Columns("H").ColumnWidth = 2.5
        
    'Format widths in output
        Columns("I").ColumnWidth = 10
        Columns("J").ColumnWidth = 18
        Columns("K").ColumnWidth = 18
        Columns("L").ColumnWidth = 18

      Range("I" & 1).Value = "Ticker"
        Range("J" & 1).Value = "Yearly Change"
        Range("K" & 1).Value = "Percent Change"
        Range("L" & 1).Value = "Total Volume"
    
    For i = 2 To LastRow
    
    'Format for color in gap
        If Range("G" & Ti).Value >= 0 Then
            Range("H" & Ti).Interior.ColorIndex = 15
        End If
  
    ' #4. output cumulative sum of volumes as Volume
        'define a Volume & accumulate to last row in each Ticker
        Volume = Volume + Range("G" & i).Value
    Range("L" & Ti).Value = Volume
    
    ' #1. output ticker
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        'Define Ti value as increase by 1
        Range("I" & Ti).Value = Range("A" & i).Value
    
    ' #2. output (last day-close value - first day-open value)
        'Define close value per ticker
        closeVal = Range("F" & i).Value
    Range("J" & Ti) = (closeVal - openVal)
        'Format for color
            If Range("J" & Ti).Value = "Yearly Change" Then
                Range("J" & Ti).Interior.ColorIndex = 2
            
            ElseIf Range("J" & Ti).Value >= 0 Then
                Range("J" & Ti).Interior.ColorIndex = 4
            
            Else
                Range("J" & Ti).Interior.ColorIndex = 3
            End If
       
    ' #3. output percentage "yearly change" / closeVal
        'format range for number style "percentage"
        Range("K" & Ti).NumberFormat = "0.00%"
    Range("K" & Ti) = ((closeVal - openVal) / openVal)
    
    ' redefine values for new start loop
    openVal = Cells(i + 1, 3)
    Volume = 0
    Ti = Ti + 1
        
        End If
    'loop end - go to next in series
    Next i
    
    'BONUS
    'How to include results of all sheets
        'locate biggest winner
        'locate biggest loser
        'locate biggest volume
End Sub
