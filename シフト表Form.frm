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
    ' マスタから情報を取得し、フォームのコンボボックスに設定する。
    Dim Lastrow As Integer  ' マスタの最終行
    
    NameBox.Clear
        Lastrow = Worksheets("助手マスタ").Cells(Rows.Count, 3).End(xlUp).Row
    For i = 4 To Lastrow
        NameBox.AddItem Worksheets("助手マスタ").Range("C" & i).Value
    Next i


    ' 今月の末日を取得する。
    Dim d As Date           ' 日
    Dim sLastday As Date    ' 今月の末日
    
    d = Range("I2").Value
        
    sLastday = Format(DateSerial(year(d), Month(d) + 1, 0), "d")
    
    
    ' 末日の日にちによってチェックボックスを非表示にする。
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

' 氏名が選択されたら処理
Private Sub NameBox_Change()
    ' 氏名コンボボックスの内容が変更された時、変更された内容を元に該当するセルから値を取得し、31個のチェックボックスの内容を変更する。
    Dim Index As Integer    ' コンボボックスの番号
    
    Index = NameBox.ListIndex

    If Index <> -1 Then
            Index = Index + 10
        
        For i = 1 To 31
            If Cells(6 + Index, 8 + i).Value = "希" Then
                Me.Controls("CheckBox" & i) = True
            Else
                Me.Controls("CheckBox" & i) = False
            End If
        Next i
    End If
End Sub

' 全て選択ボタンを押して全てのコンボボックスを選択状態にする
Private Sub SelectionButton_Click()
    ' 全て選択ボタンを押した時、全てのコンボボックスを選択状態にする処理を行う。
    For i = 1 To 31
        Me.Controls("CheckBox" & i) = True
    Next
End Sub

' 全て解除ボタンを押して全てのコンボボックスの選択を解除する
Private Sub ReleaseButton_Click()
    ' 全て解除ボタンを押した時、全てのコンボボックスの選択を解除する処理を行う。
    For i = 1 To 31
        Me.Controls("CheckBox" & i) = False
    Next
End Sub

' 完了ボタンクリック
Private Sub CompleteButton_Click()
    ' 氏名コンボボックスの内容を取得する。コンボボックスが選択されていない場合はメッセージを表示する。
    ' コンボボックスが選択されている場合はシートから年月日を取得し変数に格納し、先月のシート名を設定する。
    Dim Index As Integer    ' コンボボックスのインデックス
    Dim d As Date               ' 日
    Dim LastMonth As Date       ' 先月
    Dim y As Integer            ' 年
    Dim m As Integer            ' 月
    Dim LastSheet As String     ' 先月のシート名
    
    Index = NameBox.ListIndex
    
    If Index = -1 Then
        MsgBox "氏名が選択されていません"
    Else
        Index = Index + 10

        d = Range("I2").Value

        LastMonth = DateAdd("m", -1, d)
    
        y = year(LastMonth)
        m = Month(LastMonth)
        
        LastSheet = y & "." & m
        
        
        ' Excelの関数となる文字列を設定する。
        ' 関数は先月のシートが存在する場合にその末日を取得する関数である。
        Dim s As String ' 先月の末日文字列
        
        s = "IF(DAY(EOMONTH(" & CStr(LastSheet) & "!I3,0))=28," & CStr(LastSheet) & "!AJ" & 6 + Index & ",IF(DAY(EOMONTH(" & CStr(LastSheet) & "!I3,0))=29," & CStr(LastSheet) & "!AK" & 6 + Index & ",IF(DAY(EOMONTH(" & CStr(LastSheet) & "!I3,0))=30," & CStr(LastSheet) & "!AL" & 6 + Index & "," & CStr(LastSheet) & "!AM" & 6 + Index & ")))"
    
    
        ' 31日分処理を繰り返し、チェックボックスが選択されているかによって処理を分岐させる。
        ' チェックボックスが選択されている場合、28日以前と以降で異なる関数をセルに設定する。
        ' チェックボックスが選択されていない場合、1日、28日以降で処理を分岐させる。1日の場合、先月のシートが存在するかを確認し存在するかどうかによって異なる関数をセルに設定する。
        '28日以降の場合も同様に関数をセルに設定する。
        Dim SameSheet As Boolean    ' 同名シート判定フラグ
        
        For i = 1 To 31
            If Me.Controls("CheckBox" & i) = True Then
                If i > 28 Then
                    Cells(6 + Index, 8 + i).Value = "=IF(" & Cells(3, 8 + i).address & "="""","""",""希"")"
                Else
                    Cells(6 + Index, 8 + i).Value = "希"
                End If
            ElseIf Me.Controls("CheckBox" & i) = False Then
                If i = 1 Then
                    SameSheet = exist_check(LastSheet, "ws")
                    
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
        Next i
    End If
End Sub

' 閉じるボタン
Private Sub 閉じる_Click()
    ' フォームを閉じる。
    Unload Me
End Sub

