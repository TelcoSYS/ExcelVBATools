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
    
    CleanString = Replace(CleanString, "Á", "A")
    CleanString = Replace(CleanString, "É", "E")
    CleanString = Replace(CleanString, "Í", "I")
    CleanString = Replace(CleanString, "Ó", "O")
    CleanString = Replace(CleanString, "Ú", "U")
    CleanString = Replace(CleanString, "Ü", "U")
    CleanString = Replace(CleanString, "Ñ", "#")
    
    CleanString = Replace(CleanString, "°", ".")
    CleanString = Replace(CleanString, "ª", ".")
    CleanString = Replace(CleanString, "~", "-")
    CleanString = Replace(CleanString, Chr(150), "-")
   
    CleanString = Replace(CleanString, "Ö", "O")
    CleanString = Replace(CleanString, "Ç", "C")
    CleanString = Replace(CleanString, "Æ", "AE")
    CleanString = Replace(CleanString, "Ç", "C")
    
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