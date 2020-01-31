VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' メンバ変数を定義する。
Private adoCon As ADODB.Connection  ' ADOコネクション
Private adoRs As ADODB.Recordset    ' ADOレコードセット

' データベースへ接続
Public Function connect()
    ' データベースに接続を行う。
    ' 接続に成功した場合、Trueを返す。
    Set adoCon = New ADODB.Connection   ' ADOコネクションを作成

    On Error GoTo ErrorTrap
    
    adoCon.Open _
        "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
        " SERVER=10.206.0.47;" & _
        " DATABASE=arisan_db;" & _
        " UID=test;" & _
        " PWD=Test123_pass;"
    
    adoCon.CursorLocation = 3
Exit Function

ErrorTrap:
    ' 接続に失敗した場合、解放処理を行いメッセージを表示させる。
    Set adoCon = Nothing
    
    Select Case Err.Number
        Case -2147467259
            MsgBox "データベースに接続できません"
    End Select
End Function

' データベースへの接続を切断
Public Sub disconnect()
    ' ADOコネクションを閉じてメモリの解放を行うことでデータベースへの接続を切断する。
    adoCon.Close
    Set adoCon = Nothing
End Sub

' SQL文実行
Public Function execute(SQL As String) As ADODB.Recordset
    ' SQL文を実行し、レコードセットを戻り値として返す。
    Set adoRs = New ADODB.Recordset ' ADOレコードセットを作成

    adoRs.Open SQL, adoCon
    
    Set execute = adoRs
End Function
