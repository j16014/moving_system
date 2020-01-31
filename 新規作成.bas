Attribute VB_Name = "新規作成"
Option Explicit

' 追加
Sub NewINSERT_Click()
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Dim db As DBManager     ' データベースクラス
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    
    
    ' 確認ボタンを表示し、Yesの場合、文字数が制限値を超えていないかチェックする。
    Dim result As Long  ' YesNoボタン表示
    
    result = MsgBox("データを追加してもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
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
            
            
            ' Date関数によって今日の日付を取得し、セルの日付が今日の日付以降だった場合、年を加算する。
            Dim thisyear As Integer     ' 年
            Dim thismonth As Integer    ' 月
            Dim thisday As Integer      ' 日
            
            thisyear = year(Date)
            thismonth = Month(Date)
            thisday = Day(Date)
            
            If thismonth >= Range("B9").Value Then
                If thisday >= Range("J9").Value Then
                    thisyear = year(Date) + 1
                End If
            End If
            
            
            ' セルに日付が入力されていない場合はデフォルト値を設定する。入力されている場合は、セルの値を取得して変数に格納する。
            Dim move_day As String      ' 引越日
            Dim reception_day As String ' 受付日
            Dim preview_day As String   ' 下見日
            
            move_day = "1900-01-01"
            reception_day = "1900-01-01 01:01:00"
            preview_day = "1900-01-01 01:01:00"
            
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
            
            
            ' お客様シートのセルの値を取得し、INSERTするSQL文を設定して実行する。
            Dim SQL As String   ' SQL
            
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
            
            db.execute SQL
         

            ' データベースオブジェクトの解放処理を行う。
            db.disconnect
            Set db = Nothing
            
            
            ' セルの内容をクリアし、お客様情報シートに遷移した後、新規作成シートを非表示にする。
            Call クリア_Click
            
            Worksheets("お客様情報").Activate
            Worksheets("新規作成").Visible = False
        Else
            ' 文字数が制限値を越えている場合、データベースオブジェクトの解放処理を行いメッセージを表示する。
            Set db = Nothing
            MsgBox "文字数がオーバーしています"
        End If
    End If
Exit Sub
     
ErrorTrap:
    ' エラーが発生した場合、データベースオブジェクトの解放処理を行う
    Set db = Nothing
End Sub

' 閉じる
Sub ExitINSERT_Click()
    ' 確認ボタンを表示し、Yesの場合はセルの内容をクリアしてお客様情報シートに遷移する。そして新規作成シートを非表示にする。
    Dim result As Long  ' YesNoボタン表示

    result = MsgBox("新規作成を閉じてもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        Call クリア_Click
        
        Range("A1").Select
        Worksheets("お客様情報").Activate
        Worksheets("新規作成").Visible = False
    End If
End Sub

