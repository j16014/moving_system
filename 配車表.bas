Attribute VB_Name = "配車表"
Option Explicit

' 変数
Dim i As Integer
Dim j As Integer
Dim k As Integer

' 予約参照
Sub Re_SELECT_Click()
    ' セルの内容をクリアする。
    Range("C4:L55").Value = ""
    Range("P4:Y55").Value = ""


    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う。
    Dim db As DBManager     ' データベースクラス
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    
    
    ' レコードが存在しなかった場合に加算していくフラグを定義する。
    Dim Eofnum As Integer   ' データ無し判定フラグ
    Eofnum = 0
    

    ' 引越時間帯種別の種別数分処理を繰り返し、その都度meridianに異なる種別を設定する。
    Dim meridian As String  ' 引越時間帯種別（am,pm,free）

    For i = 1 To 3
        Select Case i
            Case 1
                meridian = "AM"
            Case 2
                meridian = "PM"
            Case 3
                meridian = "free"
        End Select
        
        
        ' SQL文を実行してレコードセットからフィールド数とレコード数を取得する。
        Dim SQL As String           ' SQL
        Dim adoRs As Object         ' ADOレコードセット
        Dim FldCount As Integer     ' フィールド数
        Dim RecCount As Long        ' レコード数

        SQL = "SELECT id,name,meridian,now_address,now_floors,now_ev,now_type," & _
        "new_address,new_floors,new_ev,new_type,preview_name,point," & _
        "start_time1,start_time2,start_time3,plan,difficulty,truck,driver," & _
        "assistant1,assistant2,assistant3,assistant4 FROM customers WHERE " & _
        "move_day = '" & Range("J1").Value & "-" & Range("M1").Value & "-" & Range("Q1").Value & "' " & _
        "AND meridian = '" & meridian & "'"
        
        Set adoRs = db.execute(SQL)

        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
            

        ' レコードが存在しない場合はフラグを加算する。
        ' レコードが存在する場合は、レコードセットを配列に格納する。
        Dim myArray() As Variant    ' 参照レコード配列
        
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
            ReDim myArray(FldCount - 1, RecCount - 1)
            myArray = adoRs.GetRows
                           
                    
            ' レコードの建物情報・住所の文字列を設定するため、レコード数分処理を繰り返す。
            Dim condition As String     ' 現・新建物情報
            
            For j = 0 To RecCount - 1
                condition = ""
                            
                   
                ' 各変数に配列から該当する値を取得し、荷物を積む地点の建物情報を設定する。
                Dim now_floors As String    ' 現階層
                Dim now_ev As String        ' 現ev
                Dim now_type As String      ' 現建物種別
                Dim flg As Boolean          ' エレベータフラグ
                            
                now_floors = myArray(4, j)
                now_ev = myArray(5, j)
                now_type = myArray(6, j)
                flg = False
                           
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
                        
                If flg = True Then
                    If now_ev = "EV有" Then
                        condition = condition & "○～"
                    Else
                        condition = condition & "×～"
                    End If
                Else
                    condition = condition & "～"
                End If
                    
                    
                ' 各変数に配列から該当する値を取得し、荷物を降ろす地点の建物情報を設定する。
                Dim new_floors As String    ' 新階層
                Dim new_ev As String        ' 新ev
                Dim new_type As String      ' 新建物種別
                
                new_floors = myArray(8, j)
                new_ev = myArray(9, j)
                new_type = myArray(10, j)
                flg = False
               
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
                        
                If flg = True Then
                    If new_ev = "EV有" Then
                        condition = condition & "○"
                    Else
                        condition = condition & "×"
                    End If
                End If
    
                
                ' 各変数に配列から該当する値を取得し、住所を設定する。
                Dim address As String       ' 現・新住所
                Dim now_address As String   ' 現住所
                Dim new_address As String   ' 新住所
                
                address = ""
                        
                now_address = myArray(3, j)
                new_address = myArray(7, j)
                address = now_address & " ～ " & new_address
                
                
                ' 配列の内容をセルに格納する。引越時間帯種別によってセルの番地が変わるため条件分岐している。
                If i = 1 Then
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
            Next j
        End If
    Next i
    
    
    ' レコードが存在しない場合はメッセージ表示する。
    If Eofnum = 3 Then
        MsgBox "お客様データがありません"
    End If
    
    
    ' レコードセットとデータベースオブジェクトの解放処理を行う。
    adoRs.Close
    Set adoRs = Nothing
    db.disconnect
    Set db = Nothing
Exit Sub
 
ErrorTrap:
    ' エラーが発生した場合、レコードセットとデータベースオブジェクトの解放処理を行う。
    Set adoRs = Nothing
    Set db = Nothing
End Sub

' 予約更新
Sub Re_UPDATA_Click()
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Dim db As DBManager     ' データベースクラス
    
    On Error GoTo ErrorTrap
   
    Set db = New DBManager
    db.connect


    ' 確認ボタンを表示しYesの場合は、配車表シートのセルの値をSQL文を設定しUPDATEを実行する。シートは二列になっているため条件分岐することで番地を変えている。
    Dim result As Long      ' YesNoボタン表示
    Dim SQL As String       ' SQL

    result = MsgBox("上書き保存してもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        For i = 0 To 1
            For j = 0 To 13
                If i = 0 Then
                    If Range("E" & j * 4 + 4).Value <> "" Then
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

                        db.execute SQL
                    End If
                Else
                    If Range("R" & j * 4 + 4).Value <> "" Then
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
                        
                        db.execute SQL
                    End If
                End If
            Next j
        Next i
    End If
 
 
    ' データベースオブジェクトの解放処理を行う。
    db.disconnect
    Set db = Nothing
Exit Sub
 
ErrorTrap:
    ' エラーが発生した場合、データベースオブジェクトの解放処理を行う。
    Set db = Nothing
End Sub
