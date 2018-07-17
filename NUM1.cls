VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Make_All_Work_This_Sheet()

numberOfBuffs = ActiveSheet.Range("B29")

Set initiativeName = ActiveSheet.Range("A6")

'edw ;exoume to prwto keli apo to row initiate'
Set ancorInitiative = ActiveSheet.Range("F6")

'edw exoume to Basico group gia tis praxeis, exodos'
Set FirstGroup = ActiveSheet.Range("F6", "AE25")

'edw exoume ta types'
Set Type1Ancor = ActiveSheet.Range("C31")
Set Type2Ancor = ActiveSheet.Range("I32")
Set Type3Ancor = ActiveSheet.Range("N32")


    Z = 0
    x = 0
    y = 1
            
    For Each cell In FirstGroup
    
    For countVal = 1 To numberOfBuffs
    
    
    'Z = cell.Column
    'h = ActiveCell.Row
    'a = Type1Ancor.Offset(y, 0)
    Set a = ActiveCell
    
    
    
    
        
    If Type1Ancor.Offset(y, 0).Value = a.Offset(0, -h).Value And Type1Ancor.Offset(y, 1).Value = a.Offset(4 - Z, 0).Value Then
        If a.Value < Type1Ancor.Offset(y, 4) Then
        a.Value = Type1Ancor.Offset(y, 4).Value
        End If
        
    End If
    
    
    y = y + 1
   
      Next countVal
            
    
    Next cell
    
    
   'Set A1 = ActiveSheet.Range("A1")
   
   'A1.Value = numberofBuffs
    


End Sub

