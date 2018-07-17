Sub Make_All_Work_This_Sheet()

numberOfBuffs = ActiveSheet.Range("B29")

Set initiativeName = ActiveSheet.Range("A6")

'edw ;exoume to prwto keli apo to row initiate'
Set ancorInitiative = ActiveSheet.Range("F6")

'edw exoume to Basico group gia tis praxeis, exodos'
Set FirstGroup = ActiveSheet.Range("F6", "AE25")

'edw exoume ta types'
Set Type1Ancor = ActiveSheet.Range("C32")
Set Type2Ancor = ActiveSheet.Range("I32")
Set Type3Ancor = ActiveSheet.Range("N32")

'loop as long anchor cell down is not blank'
'While Type1Ancor.Offset(1, 0) <> ""

    x = 0
    y = 0
    Z = 0
    
'For countVal = 1 To 28

'x = x + 1

'for countVal = 1 To 20

'y=y+1 cell.column  , cell.row

'For countVal = 1 To numberofBuffs

    'initiative'
     '   If Type1Ancor.Offset(x, y) = Type1Ancor.Offset(x, y) And Type1Ancor.Offset(x + 1, y) = 0 Then
    
   ' ancorInitiative = a
   
   x = 0
   Set a = ActiveSheet.Range("A1")
   a = 5
   a = Type1Ancor.Offset(4, y).Value
    
    For Each cell In FirstGroup
    
    For countVal = 1 To numberOfBuffs
        
    If Type1Ancor.Offset(0, y) = cell.Offset(-cell.Column, 0) And Type1Ancor.Offset(1, y) = cell.Offset(0, 4 - cell.Row) Then
        If cell.Value < Type1Ancor.Offset(4, y) Then
        cell.Value = Type1Ancor.Offset(4, y)
        End If
        
    End If
    
    
    y = y + 1
   
      Next countVal
            
    
    Next cell
    
    
   'Set A1 = ActiveSheet.Range("A1")
   
   'A1.Value = numberofBuffs
    
End Sub
