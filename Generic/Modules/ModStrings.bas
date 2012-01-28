Attribute VB_Name = "ModStrings"
''''''''''''''''''''''''''''''''''''''''''''
''            VBA Excel Tools             ''
''            Gabriel CUGLIARI            ''
''               Abr 2011                 ''
''''''''''''''''''''''''''''''''''''''''''''

''=== CleanString ===
'
'
Public Function CleanString(Cadena As String) As String
 
    
    CleanString = Trim(Cadena)
 
    Do While InStr(1, CleanString, "  ")
        CleanString = Replace(CleanString, "  ", " ")
    Loop
    
    CleanString = UCase(CleanString)
    
    CleanString = Replace(CleanString, "�", "A")
    CleanString = Replace(CleanString, "�", "E")
    CleanString = Replace(CleanString, "�", "I")
    CleanString = Replace(CleanString, "�", "O")
    CleanString = Replace(CleanString, "�", "U")
    CleanString = Replace(CleanString, "�", "U")
    CleanString = Replace(CleanString, "�", "#")
    
    CleanString = Replace(CleanString, "�", ".")
    CleanString = Replace(CleanString, "�", ".")
    CleanString = Replace(CleanString, "~", "-")
    CleanString = Replace(CleanString, Chr(150), "-")
   
    CleanString = Replace(CleanString, "�", "O")
    CleanString = Replace(CleanString, "�", "C")
    CleanString = Replace(CleanString, "�", "AE")
    CleanString = Replace(CleanString, "�", "C")
    
End Function

''=== ExtractUntil ===

Public Function ExtractUntil (txt As String, char As String)
 
    Dim ii As Integer
    
    ii = InStr(1, txt, char)
    If (ii > 0) Then
      ExtraerHasta = Left(txt, ii - 1)
    Else
      ExtraerHasta = txt
    End If
    
End Function