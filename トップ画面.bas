Attribute VB_Name = "トップ画面"
Option Explicit

Sub シフト表遷移_Click()
    ' トップ画面の右のシートに遷移
    ActiveSheet.Next.Activate
    Range("A1").Select
End Sub

Sub お客様情報遷移_Click()
    Worksheets("お客様情報").Activate
    Range("A1").Select
End Sub

Sub 配車表遷移_Click()
    Worksheets("配車表").Activate
    Range("A1").Select
End Sub

Sub 予約状況遷移_Click()
    Worksheets("予約状況").Activate
    Range("A1").Select
End Sub

