Option Explicit

' 対象：親側シート（文書管理システム出力）のシートモジュール
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit

    ' 標準モジュール側の共通処理を呼ぶ
    HandleMediaTypeCells Me, Target

SafeExit:
    ' ここでは特に何もしない（共通処理側でイベント復帰）

End Sub
