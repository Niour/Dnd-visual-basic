VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub BUFFS_UPDATE()

numberOfBuffs = ActiveSheet.Range("B33")

'edw exoume to Basico group gia tis praxeis, exodos'
Set FirstGroup = ActiveSheet.Range("R6", "AP26")
Set Competence = ActiveSheet.Range("AQ6", "AR26")

'edw exoume ta types'
Set Type1Ancor = ActiveSheet.Range("F34")
Set Type2Ancor = ActiveSheet.Range("J34")
Set Type3Ancor = ActiveSheet.Range("O34")


    x = 0
    y = 1
    
    For Each cell In Competence
    cell.Value = 0
    y = 1
    
    'type1
    For countVal = 1 To numberOfBuffs
     A = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, A)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, A)  'modifcations
    
   
    Z = Type1Ancor.Offset(y, -1).Value
    f = Type1Ancor.Offset(y, 0).Value
    g = d.Value
    h = c.Value
             
    If Type1Ancor.Offset(y, 0).Value = d.Value And Type1Ancor.Offset(y, 1).Value = e.Value Then
        cell.Value = cell.Value + Z
        
    End If
    y = y + 1
      Next countVal
    
       
    'type2
    y = 1
    
    For countVal = 1 To numberOfBuffs
    
    A = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, A)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, A)  'modifcations
    
   
    Z = Type2Ancor.Offset(y, 4).Value
    f = Type2Ancor.Offset(y, 0).Value
    g = d.Value
    h = c.Value
             
    If Type2Ancor.Offset(y, 0).Value = d.Value And Type2Ancor.Offset(y, 1).Value = e.Value Then
        cell.Value = cell.Value + Z
        
    End If
    y = y + 1
      Next countVal
            
    
    
    
    'type3
    y = 1
    
    For countVal = 1 To numberOfBuffs
    
    A = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, A)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, A)  'modifcations
    
   
    Z = Type3Ancor.Offset(y, 4).Value
    f = Type3Ancor.Offset(y, 0).Value
    g = d.Value
    h = c.Value
             
    If Type3Ancor.Offset(y, 0).Value = d.Value And Type3Ancor.Offset(y, 1).Value = e.Value Then
        cell.Value = cell.Value + Z
        
    End If
    y = y + 1
      Next countVal
    
    
    Next cell
    
    
    
    
            
    For Each cell In FirstGroup
    cell.Value = 0
    y = 1
    
    'type1
    
    For countVal = 1 To numberOfBuffs
    
    A = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, A)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, A)  'modifcations
    
   
    Z = Type1Ancor.Offset(y, -1).Value
    f = Type1Ancor.Offset(y, 0).Value
    g = d.Value
    h = c.Value
             
    If Type1Ancor.Offset(y, 0).Value = d.Value And Type1Ancor.Offset(y, 1).Value = e.Value Then
        If c.Value < Z Then
        cell.Value = Z
        End If
    End If
    y = y + 1
      Next countVal
    
   
    
    'type2
    y = 1
    
    For countVal = 1 To numberOfBuffs
    
    A = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, A)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, A)  'modifcations
    
   
    Z = Type2Ancor.Offset(y, 4).Value
    f = Type2Ancor.Offset(y, 0).Value
    g = d.Value
    h = c.Value
             
    If Type2Ancor.Offset(y, 0).Value = d.Value And Type2Ancor.Offset(y, 1).Value = e.Value Then
        If c.Value < Z Then
        cell.Value = Z
        End If
    End If
    y = y + 1
      Next countVal
            
    
    
    
    'type3
    y = 1
    
    For countVal = 1 To numberOfBuffs
    
    A = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, A)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, A)  'modifcations
    
   
    Z = Type3Ancor.Offset(y, 4).Value
    f = Type3Ancor.Offset(y, 0).Value
    g = d.Value
    h = c.Value
             
    If Type3Ancor.Offset(y, 0).Value = d.Value And Type3Ancor.Offset(y, 1).Value = e.Value Then
        If c.Value < Z Then
        cell.Value = Z
        End If
    End If
    y = y + 1
      Next countVal
            
    
    Next cell
    
End Sub


Sub DURATION_BETA()
Attribute DURATION_BETA.VB_ProcData.VB_Invoke_Func = "D\n14"

numberOfBuffs = ActiveSheet.Range("B33")


Set spellDuration = ActiveSheet.Range("C35")


x = -1


For countVal = 1 To numberOfBuffs

x = x + 1



If spellDuration.Offset(x, 0).NumberFormat <> "General" Then

spellDuration.Offset(x, 0).Value = spellDuration.Offset(x, 0).Value - 1

End If



Next countVal


End Sub


Sub DURATION()

numberOfBuffs = ActiveSheet.Range("B33")


Set spellDuration = ActiveSheet.Range("C35")


x = -1


For countVal = 1 To numberOfBuffs

x = x + 1





spellDuration.Offset(x, 0).Value = spellDuration.Offset(x, 0).Value - 1





Next countVal


End Sub

