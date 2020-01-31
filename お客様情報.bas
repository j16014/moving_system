Attribute VB_Name = "お客様情報"
Option Explicit

'お客様情報クリア
Sub クリア_Click()
    ' セルの内容をクリアする。
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
    ' フォームを表示する。
    参照Form.Show
End Sub

' 追加
Sub INSERT_Click()
    ' ワークシートを表示し移動する。
    Worksheets("新規作成").Visible = True
    Worksheets("新規作成").Activate
    Range("A1").Select
End Sub

' 更新
Sub UPDATA_Click()
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Dim db As DBManager         ' データベースクラス
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    
    
    ' 確認ボタンを表示し、確認ボタンがYesの場合、文字数が制限値を超えていないかチェックする。
    Dim result As Long  ' YesNoボタン表示
    
    result = MsgBox("上書き保存してもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)

    If result = vbYes Then
        If Range("I5").Value = "" Then
            MsgBox "IDが選択されていません"
            End
        End If
        
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
        
            
            ' Date関数によって今日の日付を取得し、セルの日付が昨日の日付以前である場合年を加算する。
            Dim thisyear As Integer     ' 年
            Dim thismonth As Integer    ' 月
            Dim thisday As Integer      ' 日
            
            thisyear = year(Date)
            thismonth = Month(Date)
            thisday = Day(Date)
            
            If thismonth >= Range("B9").Value Then
                If thisday > Range("J9").Value Then
                    thisyear = year(Date) + 1
                End If
            End If
    
    
            ' お客様情報シートのセルの値を取得してUPDATEするSQL文を設定し実行する。
            Dim SQL As String   ' SQL
            
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
            
            db.execute SQL
        End If
    End If
 

    ' データベースオブジェクトの解放処理を行う。
    db.disconnect
    Set db = Nothing
Exit Sub
 
ErrorTrap:
    ' エラーが発生した場合、データベースオブジェクトの解放処理を行う。
    Set db = Nothing
End Sub

' 削除
Sub DELETE_Click()
    ' フォームを表示する。
    削除Form.Show
End Sub
