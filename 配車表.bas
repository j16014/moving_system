Attribute VB_Name = "配車表"
Option Explicit

' 変数
Dim i As Integer
Dim j As Integer
Dim k As Integer

' 予約参照
Sub Re_SELECT_Click()
    Dim db As DBManager         ' データベースクラス
    Dim dbFlg As Boolean        ' 接続フラグ
    Dim adoRs As Object         ' ADOレコードセット
    Dim SQL As String           ' SQL
    Dim FldCount As Integer     ' フィールド数
    Dim RecCount As Long        ' レコード数
    Dim myArray() As Variant    ' 参照レコード配列
    Dim condition As String     ' 現新建物情報
    Dim address As String       ' 現新住所
    Dim flg As Boolean          ' 条件・evフラグ
    Dim Eofnum As Integer       ' データ無し判定
    Dim now_address As String   ' 現住所
    Dim now_floors As String    ' 現階層
    Dim now_ev As String        ' 現ev
    Dim now_type As String      ' 現建物種別
    Dim new_address As String   ' 新住所
    Dim new_floors As String    ' 新階層
    Dim new_ev As String        ' 新ev
    Dim new_type As String      ' 新建物種別
    Dim meridian As String      ' am,pm,free
     
    On Error GoTo ErrorTrap
   
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Set db = New DBManager
    dbFlg = db.connect
    
    ' 配車表クリア
    Range("C4:L55").Value = ""
    Range("P4:Y55").Value = ""
    
    ' データ無し判定初期化
    Eofnum = 0
    
    ' SQL文
    SQL = "SELECT id,name,meridian,now_address,now_floors,now_ev,now_type," & _
    "new_address,new_floors,new_ev,new_type,preview_name,point,start_time1,start_time2,start_time3," & _
    "plan,difficulty,truck,driver,assistant1,assistant2,assistant3,assistant4 FROM customers " & _
    "WHERE move_day = '" & Range("J1").Value & "-" & Range("M1").Value & "-" & Range("Q1").Value & "'"
    
    ' SQLの実行
    Set adoRs = db.execute(SQL)
        
    For i = 1 To 3
        ' AM・PM・freeで条件分岐
        Select Case i
            Case 1
                meridian = "AM"
            Case 2
                meridian = "PM"
            Case 3
                meridian = "free"
        End Select
        
        ' SQL文
        SQL = "SELECT id,name,meridian,now_address,now_floors,now_ev,now_type," & _
        "new_address,new_floors,new_ev,new_type,preview_name,point," & _
        "start_time1,start_time2,start_time3,plan,difficulty,truck,driver," & _
        "assistant1,assistant2,assistant3,assistant4 FROM customers WHERE " & _
        "move_day = '" & Range("J1").Value & "-" & Range("M1").Value & "-" & Range("Q1").Value & "' " & _
        "AND meridian = '" & meridian & "'"
        
        ' SQLの実行
        Set adoRs = db.execute(SQL)
            
        ' フィールド数とレコード数を取得
        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
            
        ' AM・PM・freeにレコードが無い場合加算
        If adoRs.EOF Then
            Select Case i
                Case 1
                    Eofnum = Eofnum + 1
                Case 2
                    Eofnum = Eofnum + 1
                Case 3
                    Eofnum = Eofnum + 1
            End Select
        Else
            ' 二次元配列を再定義
            ReDim myArray(FldCount - 1, RecCount - 1)
            ' レコードセットの内容を変数に格納
            myArray = adoRs.GetRows
                    
            For j = 0 To RecCount - 1
                ' 条件初期化
                condition = ""
                flg = False
                            
                ' 条件変数定義
                ' 現階層
                now_floors = myArray(4, j)
                ' 現ev
                now_ev = myArray(5, j)
                ' 現建物種別
                now_type = myArray(6, j)
                ' 新階層
                new_floors = myArray(8, j)
                ' 新ev
                new_ev = myArray(9, j)
                ' 新建物種別
                new_type = myArray(10, j)
                           
                ' 積み地の条件
                Select Case now_type
                    Case "アパート", "団地", "MC"
                        condition = condition & now_floors
                        flg = True
                    Case "社宅", "一軒家"
                        Select Case now_floors
                           Case 1
                                condition = condition & "1"
                           Case 2
                                condition = condition & "1/2"
                           Case 3
                                condition = condition & "1/2/3"
                           Case 4
                                condition = condition & "1/2/3/4"
                           Case Else
                                condition = condition & now_floors
                    End Select
                End Select
                        
                ' 積み地evと～
                If flg = True Then
                    If now_ev = "EV有" Then
                        condition = condition & "○～"
                    Else
                        condition = condition & "×～"
                    End If
                Else
                    condition = condition & "～"
                End If
                            
                ' フラグリセット
                flg = False
                        
                ' 降ろし地条件
                Select Case new_type
                    Case "アパート", "団地", "MC"
                        condition = condition & new_floors
                        flg = True
                    Case "ご新築", "社宅", "一軒家"
                        Select Case new_floors
                            Case 1
                                condition = condition & "1"
                            Case 2
                                condition = condition & "1/2"
                            Case 3
                                condition = condition & "1/2/3"
                            Case 4
                                condition = condition & "1/2/3/4"
                            Case Else
                                condition = condition & new_floors
                        End Select
                End Select
                        
                ' 降ろし地ev
                If flg = True Then
                    If new_ev = "EV有" Then
                        condition = condition & "○"
                    Else
                        condition = condition & "×"
                    End If
                End If
                        
                ' 住所初期化
                address = ""
                        
                ' 住所変数定義
                ' 現住所
                now_address = myArray(3, j)
                ' 新住所
                new_address = myArray(7, j)
                ' 現新住所
                address = now_address & " ～ " & new_address
                        
                If i = 1 Then
                    ' セルに格納
                    Range("E" & j * 4 + 4) = myArray(0, j)     ' ID
                    Range("G" & j * 4 + 4) = myArray(1, j)     ' お客様氏名
                    Range("E" & j * 4 + 6) = condition         ' 現新建物情報
                    Range("F" & j * 4 + 6) = address           ' 現新住所
                    Range("D" & j * 4 + 6) = myArray(11, j)    ' 下見担当
                    Range("F" & j * 4 + 4) = myArray(12, j)    ' ポイント数
                    Range("C" & j * 4 + 4) = myArray(13, j)    ' 開始時間1
                    Range("C" & j * 4 + 5) = myArray(14, j)    ' 開始時間2
                    Range("C" & j * 4 + 7) = myArray(15, j)    ' 開始時間3
                    Range("D" & j * 4 + 4) = myArray(16, j)    ' プラン
                    Range("I" & j * 4 + 4) = myArray(17, j)    ' 難易度
                    Range("J" & j * 4 + 4) = myArray(18, j)    ' トラック
                    Range("J" & j * 4 + 6) = myArray(19, j)    ' ドライバー
                    Range("K" & j * 4 + 4) = myArray(20, j)    ' 助手1
                    Range("K" & j * 4 + 6) = myArray(21, j)    ' 助手2
                    Range("L" & j * 4 + 4) = myArray(22, j)    ' 助手3
                    Range("L" & j * 4 + 6) = myArray(23, j)    ' 助手4
                ElseIf i = 2 Then
                    ' セルに格納
                    Range("R" & j * 4 + 4) = myArray(0, j)     ' ID
                    Range("T" & j * 4 + 4) = myArray(1, j)     ' お客様氏名
                    Range("R" & j * 4 + 6) = condition         ' 現新建物情報
                    Range("S" & j * 4 + 6) = address           ' 現新住所
                    Range("Q" & j * 4 + 6) = myArray(11, j)    ' 下見担当
                    Range("S" & j * 4 + 4) = myArray(12, j)    ' ポイント数
                    Range("P" & j * 4 + 4) = myArray(13, j)    ' 開始時間1
                    Range("P" & j * 4 + 5) = myArray(14, j)    ' 開始時間2
                    Range("P" & j * 4 + 7) = myArray(15, j)    ' 開始時間3
                    Range("Q" & j * 4 + 4) = myArray(16, j)    ' プラン
                    Range("V" & j * 4 + 4) = myArray(17, j)    ' 難易度
                    Range("W" & j * 4 + 4) = myArray(18, j)    ' トラック
                    Range("W" & j * 4 + 6) = myArray(19, j)    ' ドライバー
                    Range("X" & j * 4 + 4) = myArray(20, j)    ' 助手1
                    Range("X" & j * 4 + 6) = myArray(21, j)    ' 助手2
                    Range("Y" & j * 4 + 4) = myArray(22, j)    ' 助手3
                    Range("Y" & j * 4 + 6) = myArray(23, j)    ' 助手4
                ElseIf i = 3 Then
                    If j < 5 Then
                        ' セルに格納
                        Range("E" & j * 4 + 36) = myArray(0, j)     ' ID
                        Range("G" & j * 4 + 36) = myArray(1, j)     ' お客様氏名
                        Range("E" & j * 4 + 38) = condition         ' 現新建物情報
                        Range("F" & j * 4 + 38) = address           ' 現新住所
                        Range("D" & j * 4 + 38) = myArray(11, j)    ' 下見担当
                        Range("F" & j * 4 + 36) = myArray(12, j)    ' ポイント数
                        Range("C" & j * 4 + 36) = myArray(13, j)    ' 開始時間1
                        Range("C" & j * 4 + 37) = myArray(14, j)    ' 開始時間2
                        Range("C" & j * 4 + 39) = myArray(15, j)    ' 開始時間3
                        Range("D" & j * 4 + 36) = myArray(16, j)    ' プラン
                        Range("I" & j * 4 + 36) = myArray(17, j)    ' 難易度
                        Range("J" & j * 4 + 36) = myArray(18, j)    ' トラック
                        Range("J" & j * 4 + 38) = myArray(19, j)    ' ドライバー
                        Range("K" & j * 4 + 36) = myArray(20, j)    ' 助手1
                        Range("K" & j * 4 + 38) = myArray(21, j)    ' 助手2
                        Range("L" & j * 4 + 36) = myArray(22, j)    ' 助手3
                        Range("L" & j * 4 + 38) = myArray(23, j)    ' 助手4
                    Else
                        Range("R" & j * 4 + 16) = myArray(0, j)     ' ID
                        Range("T" & j * 4 + 16) = myArray(1, j)     ' お客様氏名
                        Range("R" & j * 4 + 18) = condition         ' 現新建物情報
                        Range("S" & j * 4 + 18) = address           ' 現新住所
                        Range("Q" & j * 4 + 18) = myArray(11, j)    ' 下見担当
                        Range("S" & j * 4 + 16) = myArray(12, j)    ' ポイント数
                        Range("P" & j * 4 + 16) = myArray(13, j)    ' 開始時間1
                        Range("P" & j * 4 + 17) = myArray(14, j)    ' 開始時間2
                        Range("P" & j * 4 + 19) = myArray(15, j)    ' 開始時間3
                        Range("Q" & j * 4 + 16) = myArray(16, j)    ' プラン
                        Range("V" & j * 4 + 16) = myArray(17, j)    ' 難易度
                        Range("W" & j * 4 + 16) = myArray(18, j)    ' トラック
                        Range("W" & j * 4 + 18) = myArray(19, j)    ' ドライバー
                        Range("X" & j * 4 + 16) = myArray(20, j)    ' 助手1
                        Range("X" & j * 4 + 18) = myArray(21, j)    ' 助手2
                        Range("Y" & j * 4 + 16) = myArray(22, j)    ' 助手3
                        Range("Y" & j * 4 + 18) = myArray(23, j)    ' 助手4
                    End If
                End If
            Next
        End If
    Next
    
    ' AM・PM・freeのデータが無い場合メッセージ表示
    If Eofnum = 3 Then
        MsgBox "お客様データがありません"
    End If
    
    ' 解放処理
    adoRs.Close
    Set adoRs = Nothing
    db.disconnect
    Set db = Nothing
Exit Sub
 
ErrorTrap:
    ' 解放処理
    Set adoRs = Nothing
    Set db = Nothing
    
    ' エラー処理
    Select Case Err.Number
        ' DB接続エラー
        Case -2147467259
            MsgBox "データベースに接続できません"
    End Select
End Sub

' 予約更新
Sub Re_UPDATA_Click()
    Dim db As DBManager     ' データベースクラス
    Dim dbFlg As Boolean    ' 接続フラグ
    Dim SQL As String       ' SQL
    Dim result As Long      ' YesNoボタン表示
    
    On Error GoTo ErrorTrap
   
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Set db = New DBManager
    dbFlg = db.connect
          
    result = MsgBox("上書き保存してもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        ' AMかPM
        For i = 0 To 1
            ' 13行（配車表の午前+free・午後+free）
            For j = 0 To 13
                ' AM
                If i = 0 Then
                    If Range("E" & j * 4 + 4).Value <> "" Then
                        ' SQL文
                        SQL = "UPDATE customers SET start_time1 = '" & Range("C" & j * 4 + 4).Value & "'," & _
                        "start_time2 = '" & Range("C" & j * 4 + 5).Value & "'," & _
                        "start_time3 = '" & Range("C" & j * 4 + 7).Value & "'," & _
                        "plan = '" & Range("D" & j * 4 + 4).Value & "'," & _
                        "difficulty = '" & Range("I" & j * 4 + 4).Value & "'," & _
                        "truck = '" & Range("J" & j * 4 + 4).Value & "'," & _
                        "driver = '" & Range("J" & j * 4 + 6).Value & "'," & _
                        "assistant1 = '" & Range("K" & j * 4 + 4).Value & "'," & _
                        "assistant2 = '" & Range("K" & j * 4 + 6).Value & "'," & _
                        "assistant3 = '" & Range("L" & j * 4 + 4).Value & "'," & _
                        "assistant4 = '" & Range("L" & j * 4 + 6).Value & "'" & _
                        " WHERE id = '" & Range("E" & j * 4 + 4).Value & "'"
                        
                        ' SQLの実行
                        db.execute SQL
                    End If
                ' PM
                Else
                    If Range("Q" & j * 4 + 4).Value <> "" Then
                        ' SQL文
                        SQL = "UPDATE customers SET start_time1 = '" & Range("P" & j * 4 + 4).Value & "'," & _
                        "start_time2 = '" & Range("P" & j * 4 + 5).Value & "'," & _
                        "start_time3 = '" & Range("P" & j * 4 + 7).Value & "'," & _
                        "plan = '" & Range("Q" & j * 4 + 4).Value & "'," & _
                        "difficulty = '" & Range("V" & j * 4 + 4).Value & "'," & _
                        "truck = '" & Range("W" & j * 4 + 4).Value & "'," & _
                        "driver = '" & Range("W" & j * 4 + 6).Value & "'," & _
                        "assistant1 = '" & Range("X" & j * 4 + 4).Value & "'," & _
                        "assistant2 = '" & Range("X" & j * 4 + 6).Value & "'," & _
                        "assistant3 = '" & Range("Y" & j * 4 + 4).Value & "'," & _
                        "assistant4 = '" & Range("Y" & j * 4 + 6).Value & "'" & _
                        " WHERE id = '" & Range("R" & j * 4 + 4).Value & "'"
                        
                        ' SQLの実行
                        db.execute SQL
                    End If
                End If
            Next j
        Next i
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
        ' DB接続エラー
        Case -2147467259
            MsgBox "データベースに接続できません"
    End Select
End Sub
