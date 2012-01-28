Attribute VB_Name = "ModInsert"
''''''''''''''''''''''''''''''''''''''''''''
''            VBA Excel Tools             ''
''            Gabriel CUGLIARI            ''
''               Abr 2011                 ''
''''''''''''''''''''''''''''''''''''''''''''


''=== genInsertSentence ===
'
Public Function genInsertSentence(ByRef Rng As Range, Optional head As String = "", Optional tail As String = "") As String

  Dim cc As Integer
  Dim Line As String

  If Rng.Rows.Count <> 1 Then
    genInsertSentence = "Invalid Row Count"
    Exit Function
  End If

  Line = ""
  For cc = 1 To Rng.Columns.Count

    Line = Line & IIf(Len(Line) = 0, "(", ", ")
    xx = VarType(Rng(1, cc))
    Select Case VarType(Rng(1, cc))
      Case 8: ''String
        Line = Line & "'" & Trim(Rng(1, cc)) & "'"
      Case 5: ''Number
        Line = Line & Trim(Rng(1, cc))
      Case 7: ''Date
        Line = Line & "'" & Format(Rng(1, cc), "yyyy-mm-dd") & "'"
      Case 0: ''Celda vacia
        Line = Line & "NULL"
      Case Else:
        Line = Line & "'" & Trim(Rng(1, cc)) & "'"
    End Select
  Next cc

  Line = IIf(Len(head) > 0, head & " ", "") & Line & ")" & tail
  genInsertSentence = Line

End Function
