Attribute VB_Name = "お客様情報"
Option Explicit

'お客様情報クリア
Sub クリア_Click()
    Range("I5") = ""    ' お客様ID
    Range("X9") = ""    ' お客様氏名
    Range("B9") = ""    ' 希望日1
    Range("J9") = ""    ' 希望日2
    Range("Q9") = ""    ' am,pm,free
    Range("S9") = ""    ' 開始時間前
    Range("V9") = ""    ' 開始時間後
    Range("I6") = ""    ' 希望日理由
    Range("AE6") = ""   ' 自宅電話番号1
    Range("AI6") = ""   ' 自宅電話番号2
    Range("AN6") = ""   ' 自宅電話番号3
    Range("AE7") = ""   ' 連絡先電話番号1
    Range("AI7") = ""   ' 連絡先電話番号2
    Range("AN7") = ""   ' 連絡先電話番号3
    Range("K12") = ""   ' 現住所
    Range("K11") = ""   ' 現〒1
    Range("O11") = ""   ' 現〒2
    Range("C13") = ""   ' 現階数
    Range("I13") = ""   ' 現ev
    Range("G14") = ""   ' 現道幅
    Range("AM11") = ""  ' 現建物種別
    Range("K17") = ""   ' 新住所
    Range("K16") = ""   ' 新〒1
    Range("O16") = ""   ' 新〒2
    Range("C18") = ""   ' 新階層
    Range("I18") = ""   ' 新ev
    Range("G19") = ""   ' 新道幅
    Range("AM16") = ""  ' 新建物種別
    Range("AR8") = ""   ' 受付日1
    Range("AV8") = ""   ' 受付日2
    Range("AZ8") = ""   ' 受付日3
    Range("BD8") = ""   ' 受付日4
    Range("AU11") = ""  ' 受付担当者
    Range("AR15") = ""  ' 下見日1
    Range("AV15") = ""  ' 下見日2
    Range("AZ15") = ""  ' 下見日3
    Range("BD15") = ""  ' 下見日4
    Range("AU18") = ""  ' 下見担当者
    Range("M21:M69") = ""   ' 荷物量
    Range("Z21:Z69") = ""
    Range("AM21:AM69") = ""
    Range("BC21:BC45") = ""
    Range("AY49") = ""
    Range("AY54") = ""
    Range("BC55:BC59") = ""
    Range("AZ73").Value = "=SUM(K71+X71+AK71+AZ71)" ' ポイント合計
End Sub

' 参照
Sub SELECT_Click()
    参照Form.Show
End Sub

' 追加
Sub INSERT_Click()
    ' ワークシート表示
    Worksheets("新規作成").Visible = True
    ' シート移動
    Worksheets("新規作成").Activate
    Range("A1").Select
End Sub

' 更新
Sub UPDATA_Click()
    Dim db As DBManager         ' データベースクラス
    Dim dbFlg As Boolean        ' 接続フラグ
    Dim SQL As String           ' SQL
    Dim result As Long          ' YesNoボタン表示
    Dim thisyear As Integer     ' 年
    Dim thismonth As Integer    ' 月
    Dim thisday As Integer      ' 日
     
    On Error GoTo ErrorTrap
    
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Set db = New DBManager
    dbFlg = db.connect
          
    result = MsgBox("上書き保存してもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        ' 文字数制限チェック
        If Len(Range("X9").Value) <= 20 And _
        Len(Range("B9").Value) <= 2 And Len(Range("J9").Value) <= 2 And _
        Len(Range("Q9").Value) <= 4 And Len(Range("S9").Value) <= 10 And Len(Range("V9").Value) <= 10 And _
        Len(Range("I6").Value) <= 255 And _
        Len(Range("AE6").Value) + Len(Range("AI6").Value) + Len(Range("AN6").Value) <= 13 And _
        Len(Range("AE7").Value) + Len(Range("AI7").Value) + Len(Range("AN7").Value) <= 13 And _
        Len(Range("K12").Value) <= 100 And Len(Range("K11").Value) + Len(Range("O11").Value) <= 7 And _
        Len(Range("C13").Value) <= 3 And Len(Range("I13").Value) <= 3 And _
        Len(Range("G14").Value) <= 1 And Len(Range("AM11").Value) <= 10 And _
        Len(Range("K17").Value) <= 100 And Len(Range("K16").Value) + Len(Range("O16").Value) <= 7 And _
        Len(Range("C18").Value) <= 3 And Len(Range("I18").Value) <= 3 And _
        Len(Range("G19").Value) <= 1 And Len(Range("AM16").Value) <= 10 And _
        Len(Range("AR8").Value) <= 2 And Len(Range("AV8").Value) <= 2 And _
        Len(Range("AZ8").Value) <= 2 And Len(Range("BD8").Value) <= 2 And Len(Range("AU11").Value) <= 20 And _
        Len(Range("AR15").Value) <= 2 And Len(Range("AV15").Value) <= 2 And _
        Len(Range("AZ15").Value) <= 2 And Len(Range("BD15").Value) <= 2 And Len(Range("AU18").Value) <= 20 And _
        Len(Range("AZ73").Value) <= 5 _
        Then
            ' 引越し年判定
            thisyear = year(Date)
            thismonth = Month(Date)
            thisday = Day(Date)
            
            If thismonth >= Range("B9").Value Then
                If thisday >= Range("J9").Value Then
                    thisyear = year(Date) + 1
                End If
            End If
    
            ' SQL文
            SQL = "UPDATE customers SET name = '" & Range("X9") & "'," & _
            "move_day = '" & thisyear & "-" & Range("B9").Value & "-" & Range("J9").Value & "'," & _
            "meridian = '" & Range("Q9").Value & "',front_time = '" & Range("S9").Value & "'," & _
            "back_time = '" & Range("V9").Value & "',reason = '" & Range("I6").Value & "'," & _
            "home_phone = '" & Range("AE6").Value & "," & Range("AI6").Value & "," & Range("AN6").Value & "'," & _
            "contact_phone = '" & Range("AE7").Value & "," & Range("AI7").Value & "," & Range("AN7").Value & "'," & _
            "now_address = '" & Range("K12").Value & "'," & _
            "now_postalcode = '" & Range("K11").Value & "," & Range("O11").Value & "'," & _
            "now_floors = '" & Range("C13").Value & "',now_ev = '" & Range("I13").Value & "'," & _
            "now_width = '" & Range("G14").Value & "',now_type = '" & Range("AM11").Value & "'," & _
            "new_address = '" & Range("K17").Value & "'," & _
            "new_postalcode = '" & Range("K16").Value & "," & Range("O16").Value & "'," & _
            "new_floors = '" & Range("C18").Value & "',new_ev = '" & Range("I18").Value & "'," & _
            "new_width = '" & Range("G19").Value & "',new_type = '" & Range("AM16").Value & "'," & _
            "reception_day = '1900-''" & Range("AR8").Value & "-" & Range("AV8").Value & " " & Range("AZ8").Value & "" & _
            ":" & Range("BD8").Value & "'':00',reception_name = '" & Range("AU11").Value & "'," & _
            "preview_day = '1900-''" & Range("AR15").Value & "-" & Range("AV15").Value & " " & Range("AZ15").Value & "" & _
            ":" & Range("BD15").Value & "'':00',preview_name = '" & Range("AU18").Value & "'," & _
            "point = '" & Range("AZ73").Value & "' WHERE id = " & Range("I5")
            
            ' SQLの実行
            db.execute SQL
        End If
    End If
 
    ' 解放処理
    db.disconnect
    Set db = Nothing
Exit Sub
 
ErrorTrap:
    ' 解放処理
    Set db = Nothing
    
    ' エラー処理
    Select Case Err.Number
        ' IDが空白
        Case -2147217900
            MsgBox "IDが選択されていません"
        ' DB接続エラー
        Case -2147467259
            MsgBox "データベースに接続できません"
    End Select
End Sub

' 削除
Sub DELETE_Click()
    削除Form.Show
End Sub
