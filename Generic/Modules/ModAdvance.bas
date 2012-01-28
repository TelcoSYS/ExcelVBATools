Attribute VB_Name = "ModAdvance"
''''''''''''''''''''''''''''''''''''''''''''
''            VBA Excel Tools             ''
''            Gabriel CUGLIARI            ''
''               Abr 2011                 ''
''''''''''''''''''''''''''''''''''''''''''''

''=== RangeFillPercert ===
'
Public Function RangeFillPercert(ByRef Rng As Range) As Integer

  Dim Total As Long
  Dim Acc As Double

  Acc = 0
  Total = Rng.Columns.Count * Rng.Rows.Count

  For cc = 1 To Rng.Columns.Count
    For rr = 1 To Rng.Rows.Count
      If VarType(Rng(rr, cc)) > 0 Then
        Acc = Acc + 1
      End If
    Next rr
  Next cc

  RangeFillPercert = Int((Acc * 100) / Total)

End Function
