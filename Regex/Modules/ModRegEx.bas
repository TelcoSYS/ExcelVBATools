Attribute VB_Name = "ModRegEx"
''''''''''''''''''''''''''''''''''''''''''''
'' Regular Expresions Macro Base Document ''
''            Gabriel CUGLIARI            ''
''               Abr 2011                 ''
''''''''''''''''''''''''''''''''''''''''''''
'' Tutorials:
''   http://www.regular-expressions.info/vbscript.html

Private RegExObj As Object

Private Sub RegExInit(fMultiLine As Boolean, fGlobal As Boolean, fIgnoreCase As Boolean)
    
    If RegExObj Is Nothing Then Set RegExObj = CreateObject("VBScript.RegExp")
    
    With RegExObj
        .MultiLine = fMultiLine
        .Global = fGlobal
        .IgnoreCase = fIgnoreCase
    End With
          
End Sub

''=== RegExTest ===
'=RegExTest(<InputString>;<Pattern>;<ReturnWhenMatch>;<WhenNot>)
'=RegExTest(A1;"ant";"OK";"wrong")
'
Public Function RegExTest(strData As String, Pattern As String, Macth As String, NoMatch As String) As String
    
    
    RegExInit True, False, True

    RegExObj.Pattern = Pattern
     
    If RegExObj.Test(strData) Then
        RegExTest = Macth
    Else
        RegExTest = NoMatch
    End If
     

End Function

''=== RegExReplace ===
'=RegExReplace(<InputString>;<Pattern>;<ReplaceString>)
'=RegExReplace(A1;"a(.)t";"X$1X")
'
Public Function RegExReplace(strData As String, Pattern As String, Replace As String) As String
        
    RegExInit True, False, True
    
    RegExObj.Pattern = Pattern
     
    RegExReplace = RegExObj.Replace(strData, Replace)

End Function

''=== RegExMatch ===
'=RegExMatch(<InputString>;<Pattern>;<ReturnWhenNotMatch>)
'=RegExMatch(A1;"a(.)t";"-----")
'
Public Function RegExMatch(strData As String, Pattern As String, Default As String) As String
    
    Dim RegExMatches As Object
    
    RegExInit True, False, True
    RegExMatch = Default
    On Local Error Resume Next
     
    RegExObj.Pattern = Pattern
     
    Set RegExMatches = RegExObj.Execute(strData)
    RegExMatch = RegExMatches(0).SubMatches(0)
     
End Function
