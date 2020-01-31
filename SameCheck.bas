Attribute VB_Name = "SameCheck"
Option Explicit

' 同名存在チェック
Public Function exist_check(Work As String, Unit As String) As Boolean
    ' 現在開いているワークブック・ワークシートを全て取得し、引数として渡された名前と同名のワークブック・ワークシートが存在するかを判定する。戻り値としてブール値を返す。
    Dim wb As Workbook      ' ワークブック
    Dim ws As Worksheet     ' ワークシート
    Dim SameFlg As Boolean  ' 同名フラグ
    
    SameFlg = False
    
    Select Case Unit
        Case "wb"
            For Each wb In Workbooks
                If wb.name = Work Then
                    SameFlg = True
                End If
            Next
        Case "ws"
            For Each ws In Sheets
                If ws.name = Work Then
                    SameFlg = True
                End If
            Next
    End Select

    exist_check = SameFlg
End Function
