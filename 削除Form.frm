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

' 起動時設定
Private Sub UserForm_Initialize()
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
    Dim db As DBManager     ' データベースクラス
    
    On Error GoTo ErrorTrap
   
    Set db = New DBManager
    db.connect
    
    
    ' 確認ボタンを表示し、確認ボタンがYesの場合、SQL文を実行し、レコードセットからフィールド数とレコード数を取得する。
    Dim adoRs As Object         ' ADOレコードセット
    Dim SQL As String           ' SQL
    Dim FldCount As Integer     ' フィールド数
    Dim RecCount As Long        ' レコード数

    SQL = "SELECT id,name,DATE_FORMAT(move_day,'%c月%e日'),preview_name," & _
    "DATE_FORMAT(preview_day,'%c月%e日 %H時%i分') FROM customers"
     
    Set adoRs = db.execute(SQL)
    
    FldCount = adoRs.Fields.Count
    RecCount = adoRs.RecordCount
    
    
    ' レコードが無い場合はメッセージを表示する。
    ' レコードが存在する場合はレコードセットを配列に格納し、リストボックスに登録する。リストボックスは複数選択可能にする。
    Dim myArray() As Variant    ' 参照レコード配列
    
    If adoRs.EOF Then
          MsgBox "お客様データがありません"
          End
    Else
        ReDim myArray(FldCount - 1, RecCount - 1)
        myArray = adoRs.GetRows
    
        With お客様リスト
            .Clear
            .ColumnCount = 5
            .ColumnWidths = "30;70;70;70"
            .Column = myArray
            .ListStyle = fmListStyleOption
            .MultiSelect = fmMultiSelectMulti
        End With
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

' 削除
Private Sub DELETEButton_Click()
    ' DBManagerクラスのdbをインスタンス化し、接続処理を行う。
    Dim db As DBManager     ' データベースクラス
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    

    ' 確認ボタンを表示し、Yesの場合はリストボックスで選択されている項目ごとにDELETEを実行する。
    Dim result As String        ' YesNoボタン表示
    Dim SQL As String   ' SQL
    
    result = MsgBox("データを削除してもよろしいですか？", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        With お客様リスト
            For i = 0 To .ListCount - 1
                If .Selected(i) = True Then
                    SQL = "DELETE FROM customers WHERE id=" & .List(i, 0)
                    
                    db.execute SQL
                End If
            Next i
        End With
        

        ' レコード削除後フォームに再表示を行う。
        ' SQL文を実行し、レコードセットからフィールド数とレコード数を取得する。
        Dim adoRs As Object         ' ADOレコードセット
        Dim FldCount As Integer     ' フィールド数
        Dim RecCount As Long        ' レコード数
        
        SQL = "SELECT id,name,DATE_FORMAT(move_day,'%c月%e日'),preview_name," & _
        "DATE_FORMAT(preview_day,'%c月%e日 %H時%i分') FROM customers"
         
        Set adoRs = db.execute(SQL)
        
        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
        
        
        ' レコードが存在しない場合は、リストボックスをクリアする。
        ' レコードが存在する場合はレコードセットを配列に格納し、リストボックスに登録する。リストボックスは複数選択可能にする。
        Dim myArray() As Variant    ' 参照レコード配列
        
        If adoRs.EOF Then
            With お客様リスト
                .Clear
            End With
        Else
            ReDim myArray(FldCount - 1, RecCount - 1)
            myArray = adoRs.GetRows
        
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
    
    
    ' データベースオブジェクトの解放処理を行う。
    db.disconnect
    Set db = Nothing
Exit Sub

ErrorTrap:
    ' エラーが発生した場合、データベースオブジェクトの解放処理を行う。
    Set db = Nothing
End Sub

' 閉じるボタン
Private Sub 閉じる_Click()
    ' フォームを閉じる。
    Unload Me
End Sub

