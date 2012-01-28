Attribute VB_Name = "ModPaste"
Sub PasteValues()
Attribute PasteValues.VB_ProcData.VB_Invoke_Func = "v\n14"
'
' PasteValues Macro
'

On Local Error GoTo PasteValues_err

    If Application.CutCopyMode Then
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Else
        ActiveSheet.PasteSpecial Format:="Texto", Link:=False, DisplayAsIcon:=False
        ''ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False
    End If

    ''Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ''ActiveSheet.PasteSpecial Format:="Texto", Link:=False, DisplayAsIcon:=False
PasteValues_err:
    
End Sub


''Application.OnKey "^v", "DoMyPaste"

Public Sub DoMyPaste()
    If Selection.[is marked cell] Then
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Else
        ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False
    End If
End Sub
