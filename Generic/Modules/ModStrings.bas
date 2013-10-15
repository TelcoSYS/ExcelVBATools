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

Public Function ExtractUntil (txt As String, char As String) As String
 
    Dim ii As Integer
    
    ii = InStr(1, txt, char)
    If (ii > 0) Then
      ExtractUntil = Left(txt, ii - 1)
    Else
      ExtractUntil = txt
    End If
    
End Function


''=== ExtractBetween ===

Public Function ExtractBetween(txt As String, ini As String, fin As String, dft As String) As String

    Dim ii As Integer, fi As Integer
   
    ii = InStr(1, txt, ini)
    If (ii > 0) Then
      ExtractBetween = Mid(txt, ii - 1)
    Else
      ExtractBetween = txt
    End If
   
    ii = InStr(1, txt, char)
    If (ii > 0) Then
      ExtractBetween = Left(txt, ii - 1)
    Else
      ExtractBetween = txt
    End If
   
End Function