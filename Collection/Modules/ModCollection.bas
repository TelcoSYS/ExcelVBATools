Attribute VB_Name = "ModCollection"
''''''''''''''''''''''''''''''''''''''''''''
''            VBA Excel Tools             ''
''            Gabriel CUGLIARI            ''
''               Abr 2011                 ''
''''''''''''''''''''''''''''''''''''''''''''

''=== JoinRange === 
'
'=JoinRange(<CellRange>;<Separator>)
'=JoinRange(A1:G1;";")
'
Public Function JoinRange(Rng As Range, separator As String) As String

  Dim nRow As Integer

  If Not Rng Is Nothing And Rng.Rows.Count = 1 Then
    nRow = Rng.Columns.Count
    JoinRange = ""
    For ii = 1 To nRow
      JoinRange = JoinRange & IIf(ii > 1, separator, "") & Trim(Rng(1, ii))
    Next
  Else
    JoinRange = "(null)"
  End If

End Function
