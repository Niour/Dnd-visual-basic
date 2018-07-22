Sub BUFFS_UPDATE()

numberOfBuffs = ActiveSheet.Range("B32")

'edw exoume to Basico group gia tis praxeis, exodos'
Set FirstGroup = ActiveSheet.Range("J6", "AH25")
Set Competence = ActiveSheet.Range("AI6", "AJ25")

'edw exoume ta types'
Set Type1Ancor = ActiveSheet.Range("D34")
Set Type2Ancor = ActiveSheet.Range("J34")
Set Type3Ancor = ActiveSheet.Range("O34")


    x = 0
    y = 1
    
    For Each cell In Competence
    cell.Value = 0
    y = 1
    
    'type1
    For countVal = 1 To numberOfBuffs
     a = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, a)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, a)  'modifcations
    
   
    Z = Type1Ancor.Offset(y, 4).Value
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
    
    a = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, a)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, a)  'modifcations
    
   
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
    
    a = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, a)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, a)  'modifcations
    
   
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
    
    a = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, a)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, a)  'modifcations
    
   
    Z = Type1Ancor.Offset(y, 4).Value
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
    
    a = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, a)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, a)  'modifcations
    
   
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
    
    a = cell.Column
    b = cell.Row
    Set c = ActiveSheet.Cells(b, a)
    
    
    Set d = ActiveSheet.Cells(b, 1) 'stats
    Set e = ActiveSheet.Cells(4, a)  'modifcations
    
   
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


Sub DURATION()

numberOfBuffs = ActiveSheet.Range("B32")


Set spellDuration = ActiveSheet.Range("C35")


x = -1

For countVal = 1 To numberOfBuffs

x = x + 1

spellDuration.Offset(x, 0).Value = spellDuration.Offset(x, 0).Value - 1

Next countVal


End Sub
