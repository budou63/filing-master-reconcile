Option Explicit

' 対象：Mokuroku シートのシートモジュールに貼り付け
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit

    ' 標準モジュール側の共通処理を呼ぶ
    HandleMediaTypeCells Me, Target

SafeExit:
    ' ここでは特に何もしない（共通処理側でイベント復帰）

End Sub
