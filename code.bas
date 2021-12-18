Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)

On Error GoTo errorhandler

    ActiveSheet.Name = ActiveSheet.Range("A1")
    
    Exit Sub
    
errorhandler:

    ActiveSheet.Range("A1") = Trim(Application.InputBox("シート名を空欄にすることはできません。", Title:="エラー"))

End Sub
