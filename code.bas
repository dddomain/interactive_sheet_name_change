Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)

On Error GoTo errorhandler
    
    ActiveSheet.Name = ActiveSheet.Range("A1")
    
    Exit Sub
    
errorhandler:

    MsgBox "シート名は変更されません。"

End Sub
