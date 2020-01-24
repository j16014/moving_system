VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} シフト表Form 
   Caption         =   "希望休入力"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4005
   OleObjectBlob   =   "シフト表Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "シフト表Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 変数
Dim i As Integer
Dim j As Integer
Dim k As Integer

' 起動時設定
Private Sub UserForm_Initialize()
    Dim d As Date           ' シートの日時変数
    Dim sLast As Date       ' 翌月1日の前日
    Dim sLastday As Date    ' 今月の末日
    Dim Lastrow As Integer  ' マスタの最終行
    
    ' コンボボックス初期値 マスタから取得
    NameBox.Clear
        Lastrow = Worksheets("助手マスタ").Cells(Rows.Count, 3).End(xlUp).Row
    For i = 4 To Lastrow
        NameBox.AddItem Worksheets("助手マスタ").Range("C" & i).Value
    Next

    d = Range("I2").Value
        
    ' 翌月１日の前日を取得
    sLast = DateSerial(year(d), Month(d) + 1, 0)
    
    ' 末日の日のみを取得
    sLastday = Format(sLast, "d")
    
    ' 末日の日数によってラベルとボタンを非表示
    If sLastday = 28 Then
        Label29.Enabled = False
        Label30.Enabled = False
        Label31.Enabled = False
        CheckBox29.Enabled = False
        CheckBox30.Enabled = False
        CheckBox31.Enabled = False
    ElseIf sLastday = 29 Then
        Label30.Enabled = False
        Label31.Enabled = False
        CheckBox30.Enabled = False
        CheckBox31.Enabled = False
    ElseIf sLastday = 30 Then
        Label31.Enabled = False
        CheckBox31.Enabled = False
    End If
End Sub

' 完了ボタンクリック
Private Sub CompleteButton_Click()
    Dim s As String             ' 先月の末尾文字列
    Dim str As String           ' 先月文字列
    Dim Index As Integer        ' コンボボックスの番号変数
    Dim d As Date               ' シートの日時変数
    Dim y As Integer            ' 年変数
    Dim m As Integer            ' 月変数
    Dim LastMonth As Date       ' 先月
    Dim SameSheet As Boolean    ' 同名シート判定フラグ
    Dim ws As Worksheet         ' ワークシート
    
    ' 氏名選択処理
    Index = NameBox.ListIndex
    
    ' 氏名が選択されていなければエラー、助手ならIndex + 10
    If Index = -1 Then
        MsgBox "氏名が選択されていません"
        Exit Sub
    Else
        Index = Index + 10

        ' シフト表の日時取得
        d = Range("I2").Value
        ' 月を一つ引く（先月にする）
        LastMonth = DateAdd("m", -1, d)
    
        ' 先月の年と月
        y = year(LastMonth)
        m = Month(LastMonth)
        
        ' 先月のシート名
        str = y & "." & m
        
        ' 先月の末尾ショートカットs
        s = "IF(DAY(EOMONTH(" & CStr(str) & "!I3,0))=28," & CStr(str) & "!AJ" & 6 + Index & "," & _
        "IF(DAY(EOMONTH(" & CStr(str) & "!I3,0))=29," & CStr(str) & "!AK" & 6 + Index & "," & _
        "IF(DAY(EOMONTH(" & CStr(str) & "!I3,0))=30," & CStr(str) & "!AL" & 6 + Index & "," & _
        "" & CStr(str) & "!AM" & 6 + Index & ")))"
    
        ' コンボボックスの値による分岐
        For i = 1 To 31
            ' 1日は前月の最終日を反映　28日以降が無い月は空白
            If Me.Controls("CheckBox" & i) = True Then
                If i > 28 Then
                    Cells(6 + Index, 8 + i).Value = "=IF(" & Cells(3, 8 + i).address & "="""","""",""希"")"
                Else
                    Cells(6 + Index, 8 + i).Value = "希"
                End If
            ElseIf Me.Controls("CheckBox" & i) = False Then
                If i = 1 Then
                    SameSheet = False
                    
                    ' ブック内に先月のシートが存在するか確認
                    For Each ws In Sheets
                        If ws.name = str Then
                            ' 存在する
                            SameSheet = True
                        End If
                    Next
                    
                    If SameSheet = True Then
                        Cells(6 + Index, 8 + i).Value = _
                        "=IF(OR(" & s & "=""休""," & s & "=""希""," & s & "=""AM""," & s & "=""PM""),1," & s & "+1)"
                    Else
                        Cells(6 + Index, 8 + i).Value = 1
                    End If
                ElseIf i > 28 Then
                    Cells(6 + Index, 8 + i).Value = _
                    "=IF(" & Cells(3, 8 + i).address & "="""",""""," & _
                    "IF(OR(" & Cells(6 + Index, 7 + i).address & "=""休""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""希""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""AM""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""PM""),1," & Cells(6 + Index, 7 + i).address & "+1))"
                Else
                    Cells(6 + Index, 8 + i).Value = _
                    "=IF(OR(" & Cells(6 + Index, 7 + i).address & "=""休""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""希""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""AM""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""PM""),1," & Cells(6 + Index, 7 + i).address & "+1)"
                End If
            End If
        Next
    End If
End Sub

' ボタンを押して全てのコンボボックスの選択を変更
Private Sub SelectionButton_Click()
    For i = 1 To 31
        Me.Controls("CheckBox" & i) = True
    Next
End Sub

' ボタンを押して全てのコンボボックスの選択を変更
Private Sub ReleaseButton_Click()
    For i = 1 To 31
        Me.Controls("CheckBox" & i) = False
    Next
End Sub

' 氏名が選択されたら処理
Private Sub NameBox_Change()
    Dim Index As Integer    ' コンボボックスの番号
    
    Index = NameBox.ListIndex

    ' 既に希望休が選択されている日はチェックを入れる
    If Index <> -1 Then
            Index = Index + 10
        
        For i = 1 To 31
            If Cells(6 + Index, 8 + i).Value = "希" Then
                Me.Controls("CheckBox" & i) = True
            Else
                Me.Controls("CheckBox" & i) = False
            End If
        Next
    End If
End Sub

' 閉じるボタン
Private Sub 閉じる_Click()
    Unload Me
End Sub

