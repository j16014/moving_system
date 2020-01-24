VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 削除Form 
   Caption         =   "削除"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7260
   OleObjectBlob   =   "削除Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "削除Form"
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
    
        ' リストボックスを登録（チェックボックス・複数選択可）
        With お客様リスト
            .Clear
            .ColumnCount = 5
            .ColumnWidths = "30;70;70;70"
            .Column = myArray
            .ListStyle = fmListStyleOption
            .MultiSelect = fmMultiSelectMulti
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

' 削除
Private Sub DELETEButton_Click()
    Dim db As DBManager         ' データベースクラス
    Dim dbFlg As Boolean        ' 接続フラグ
    Dim SQL As String           ' SQL
    Dim adoRs As Object         ' ADOレコードセット
    Dim FldCount As Integer     ' フィールド数
    Dim RecCount As Long        ' レコード数
    Dim myArray() As Variant    ' 参照レコード配列
    Dim result As String        ' YesNoボタン表示

    On Error GoTo ErrorTrap
    
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Set db = New DBManager
    dbFlg = db.connect
    
    result = MsgBox("データを削除してもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        ' リストボックスで選択したお客様データを削除
        With お客様リスト
            For i = 0 To .ListCount - 1
                If .Selected(i) = True Then
                    ' SQL文
                    SQL = "DELETE FROM customers WHERE id=" & .List(i, 0)
                      
                    ' SQLの実行
                    db.execute SQL
                End If
            Next i
        End With
        
        ' 削除後再表示処理
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
            With お客様リスト
                .Clear
            End With
        Else
            ' 二次元配列を再定義
            ReDim myArray(FldCount - 1, RecCount - 1)
            ' レコードセットの内容を変数に格納
            myArray = adoRs.GetRows
        
            ' リストボックスを更新（チェックボックス・複数選択可）
            With お客様リスト
                .Clear
                .ColumnCount = 5
                .ColumnWidths = "30;70;70;70"
                .Column = myArray
                .ListStyle = fmListStyleOption
                .MultiSelect = fmMultiSelectMulti
            End With
        End If

        ' 解放処理
        adoRs.Close
        Set adoRs = Nothing
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

' 閉じるボタン
Private Sub 閉じる_Click()
    Unload Me
End Sub

