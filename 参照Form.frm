VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 参照Form 
   Caption         =   "参照"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7185
   OleObjectBlob   =   "参照Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "参照Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 変数
Dim i As Integer
Dim j As Integer
Dim k As Integer

Private Sub UserForm_Initialize()
    Dim db As DBManager         ' データベースクラス
    Dim dbFlg As Boolean        ' 接続フラグ
    Dim adoRs As Object         ' ADOレコードセット
    Dim SQL As String           ' SQL
    Dim FldCount As Integer     ' フィールド数
    Dim RecCount As Long        ' レコード数
    Dim myArray() As Variant    ' 参照レコード配列
    
    For i = 1 To 5
        ' テキストボックス初期化
        Controls("TextBox" & i).Value = ""
        ' テキストボックス編集不可
        Controls("TextBox" & i).Locked = True
    Next i
    
    On Error GoTo ErrorTrap
    
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Set db = New DBManager
    dbFlg = db.connect
    
    ' SQL文
    SQL = "SELECT id,name,DATE_FORMAT(move_day,'%c月%e日'),preview_name," & _
    "DATE_FORMAT(preview_day,'%c月%e日 %H時%i分') FROM customers"
    
    ' SQLの実行
    Set adoRs = db.execute(SQL)

    ' フィールド数とレコード数を取得
    FldCount = adoRs.Fields.Count
    RecCount = adoRs.RecordCount
    
    ' レコードが無い場合
    If adoRs.EOF Then
        MsgBox "お客様データがありません"
        ' プログラムの終了
        End
    Else
        ' 二次元配列を再定義
        ReDim myArray(FldCount - 1, RecCount - 1)
        ' レコードセットの内容を変数に格納
        myArray = adoRs.GetRows
    
        ' リストボックスに表示
        With お客様リスト
            .Clear
            .ColumnCount = 5
            .ColumnWidths = "30;70;70;70"
            .Column = myArray
        End With
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

' 参照
Private Sub SELECTButton_Click()
    Dim db As DBManager         ' データベースクラス
    Dim dbFlg As Boolean        ' 接続フラグ
    Dim adoRs As Object         ' ADOレコードセット
    Dim SQL As String           ' SQL
    Dim FldCount As Integer     ' フィールド数
    Dim RecCount As Long        ' レコード数
    Dim myArray() As Variant    ' 参照レコード配列
    Dim splitArray() As String  ' 区切り配列
    
    ' セルの内容をクリア
    Call クリア_Click
    
    On Error GoTo ErrorTrap
    
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Set db = New DBManager
    dbFlg = db.connect
    
    ' SQL文
    SQL = "SELECT name,DATE_FORMAT(move_day,'%c,%e'),meridian,front_time,back_time,reason," & _
    "home_phone,contact_phone,now_address,now_postalcode,now_floors,now_ev,now_width,now_type," & _
    "new_address,new_postalcode,new_floors,new_ev,new_width,new_type," & _
    "DATE_FORMAT(reception_day,'%c,%e,%H,%i'),reception_name," & _
    "DATE_FORMAT(preview_day,'%c,%e,%H,%i'),preview_name,point " & _
    "FROM customers WHERE id = '" & TextBox1.Value & "'"

    ' SQLの実行
    Set adoRs = db.execute(SQL)
    
    ' フィールド数とレコード数を取得
    FldCount = adoRs.Fields.Count
    RecCount = adoRs.RecordCount
    
    ' レコードが無い場合
    If adoRs.EOF Then
        MsgBox "お客様データが選択されていません"
    Else
        ' 二次元配列を再定義
        ReDim myArray(FldCount - 1, RecCount - 1)
        ' レコードセットの内容を変数に格納
        myArray = adoRs.GetRows

        Range("I5") = TextBox1.Value        ' お客様ID
        Range("X9") = myArray(0, 0)         ' お客様氏名
        splitArray = Split(myArray(1, 0), ",")
        Range("B9") = splitArray(0)         ' 希望日1
        Range("J9") = splitArray(1)         ' 希望日2
        Range("Q9") = myArray(2, 0)         ' am,pm,free
        Range("S9") = myArray(3, 0)         ' 開始時間前
        Range("V9") = myArray(4, 0)         ' 開始時間後
        Range("I6") = myArray(5, 0)         ' 希望日理由
        splitArray = Split(myArray(6, 0), ",")
        Range("AE6") = splitArray(0)        ' 自宅電話番号1
        Range("AI6") = splitArray(1)        ' 自宅電話番号2
        Range("AN6") = splitArray(2)        ' 自宅電話番号3
        splitArray = Split(myArray(7, 0), ",")
        Range("AE7") = splitArray(0)        ' 連絡先電話番号1
        Range("AI7") = splitArray(1)        ' 連絡先電話番号2
        Range("AN7") = splitArray(2)        ' 連絡先電話番号3
        Range("K12") = myArray(8, 0)        ' 現住所
        splitArray = Split(myArray(9, 0), ",")
        Range("K11") = splitArray(0)        ' 現〒1
        Range("O11") = splitArray(1)        ' 現〒2
        Range("C13") = myArray(10, 0)       ' 現階数
        Range("I13") = myArray(11, 0)       ' 現ev
        Range("G14") = myArray(12, 0)       ' 現道幅
        Range("AM11") = myArray(13, 0)      ' 現建物種別
        Range("K17") = myArray(14, 0)       ' 新住所
        splitArray = Split(myArray(15, 0), ",")
        Range("K16") = splitArray(0)        ' 新〒1
        Range("O16") = splitArray(1)        ' 新〒2
        Range("C18") = myArray(16, 0)       ' 新階層
        Range("I18") = myArray(17, 0)       ' 新ev
        Range("G19") = myArray(18, 0)       ' 新道幅
        Range("AM16") = myArray(19, 0)      ' 新建物種別
        splitArray = Split(myArray(20, 0), ",")
        Range("AR8") = splitArray(0)        ' 受付日1
        Range("AV8") = splitArray(1)        ' 受付日2
        Range("AZ8") = splitArray(2)        ' 受付日3
        Range("BD8") = splitArray(3)        ' 受付日4
        Range("AU11") = myArray(21, 0)      ' 受付担当者
        splitArray = Split(myArray(22, 0), ",")
        Range("AR15") = splitArray(0)       ' 下見日1
        Range("AV15") = splitArray(1)       ' 下見日2
        Range("AZ15") = splitArray(2)       ' 下見日3
        Range("BD15") = splitArray(3)       ' 下見日4
        Range("AU18") = myArray(23, 0)      ' 下見担当者
        Range("AZ73") = "=SUM(K71+X71+AK71+AZ71)+" & myArray(24, 0) 'ポイント
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

Private Sub お客様リスト_Click()
    ' クリックしたお客様データをテキストボックスに格納
    Dim sIndex
    sIndex = お客様リスト.ListIndex
    TextBox1.Text = お客様リスト.List(sIndex, 0)
    TextBox2.Text = お客様リスト.List(sIndex, 1)
    TextBox3.Text = お客様リスト.List(sIndex, 2)
    TextBox4.Text = お客様リスト.List(sIndex, 3)
    TextBox5.Text = お客様リスト.List(sIndex, 4)
End Sub

' 閉じるボタン
Private Sub 閉じる_Click()
    Unload Me
End Sub
