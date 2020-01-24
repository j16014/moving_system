Attribute VB_Name = "新規作成"
Option Explicit

' 追加
Sub NewINSERT_Click()
    Dim db As DBManager         ' データベースクラス
    Dim dbFlg As Boolean        ' 接続フラグ
    Dim SQL As String           ' SQL
    Dim result As Long          ' YesNoボタン表示
    Dim thisyear As Integer     ' 年
    Dim thismonth As Integer    ' 月
    Dim thisday As Integer      ' 日
    Dim move_day As String      ' 引越し日
    Dim reception_day As String ' 受付日
    Dim preview_day As String   ' 下見日
      
    On Error GoTo ErrorTrap
    
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Set db = New DBManager
    dbFlg = db.connect
    
    result = MsgBox("データを追加してもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
    
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
            
            ' 受付日と下見日を初期化
            move_day = "1900-01-01"
            reception_day = "1900-01-01 01:01:00"
            preview_day = "1900-01-01 01:01:00"
            
            ' 月日時の項目が空白の場合、値を設定
            If Range("B9").Value <> "" And Range("J9").Value <> "" Then
                move_day = thisyear & "-" & Range("B9").Value & "-" & Range("J9").Value
            End If
            
            If Range("AR8").Value <> "" And Range("AV8").Value <> "" And Range("AZ8").Value <> "" And Range("BD8").Value <> "" Then
                reception_day = "1900-" & Range("AR8").Value & "-" & Range("AV8").Value & " " & _
                "" & Range("AZ8").Value & ":" & Range("BD8").Value & ":00"
            End If
            
            If Range("AR15").Value <> "" And Range("AV15").Value <> "" And Range("AZ15").Value <> "" And Range("BD15").Value <> "" Then
                preview_day = "1900-" & Range("AR15").Value & "-" & Range("AV15").Value & " " & _
                "" & Range("AZ15").Value & ":" & Range("BD15").Value & ":00"
            End If
            
            ' SQL文
            SQL = "INSERT INTO customers (name,move_day,meridian,front_time,back_time,reason,home_phone,contact_phone," & _
            "now_address,now_postalcode,now_floors,now_ev,now_width,now_type," & _
            "new_address,new_postalcode,new_floors,new_ev,new_width,new_type," & _
            "reception_day,reception_name,preview_day,preview_name,point,start_time1,start_time2,start_time3," & _
            "plan,difficulty,truck,driver,assistant1,assistant2,assistant3,assistant4) " & _
            " VALUES('" & Range("X9").Value & "','" & move_day & "'," & _
            "'" & Range("Q9").Value & "','" & Range("S9").Value & "','" & Range("V9").Value & "'," & _
            "'" & Range("I6").Value & "'," & _
            "'" & Range("AE6").Value & "," & Range("AI6").Value & "," & Range("AN6").Value & "'," & _
            "'" & Range("AE7").Value & "," & Range("AI7").Value & "," & Range("AN7").Value & "'," & _
            "'" & Range("K12").Value & "','" & Range("K11").Value & "," & Range("O11").Value & "'," & _
            "'" & Range("C13").Value & "','" & Range("I13").Value & "','" & Range("G14").Value & "'," & _
            "'" & Range("AM11").Value & "'," & _
            "'" & Range("K17").Value & "','" & Range("K16").Value & "," & Range("O16").Value & "'," & _
            "'" & Range("C18").Value & "','" & Range("I18").Value & "','" & Range("G19").Value & "'," & _
            "'" & Range("AM16").Value & "'," & _
            "'" & reception_day & "','" & Range("AU11").Value & "'," & _
            "'" & preview_day & "','" & Range("AU18").Value & "'," & _
            "'" & Range("AZ73").Value & "','-','-','-','-','-','-','-','-','-','-','-' )"
            
            ' SQLの実行
            db.execute SQL
         
            ' 解放処理
            db.disconnect
            Set db = Nothing
            
            ' セルの内容クリア
            Call クリア_Click
            
            ' シート移動
            Worksheets("お客様情報").Activate
            ' ワークシート非表示
            Worksheets("新規作成").Visible = False
        Else
            ' 解放処理
            Set db = Nothing
            MsgBox "文字数がオーバーしています"
        End If
    End If
Exit Sub
     
ErrorTrap:
    ' 解放処理
    Set db = Nothing
    
    ' エラー処理
    Select Case Err.Number
        ' DB接続エラー
        Case -2147467259
            MsgBox "データベースに接続できません"
    End Select
End Sub

' 閉じる
Sub ExitINSERT_Click()
    Dim result As Long  ' YesNoボタン表示

    result = MsgBox("新規作成を閉じてもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        ' セルの内容クリア
        Call クリア_Click
        
        Range("A1").Select
        
        ' シート移動
        Worksheets("お客様情報").Activate
        ' ワークシート非表示
        Worksheets("新規作成").Visible = False
    End If
End Sub

