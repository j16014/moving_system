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

Private adoCon As ADODB.Connection  ' ADOコネクション
Private adoRs As ADODB.Recordset    ' ADOレコードセット

' データベースへ接続
Public Function connect() As Boolean
    Set adoCon = New ADODB.Connection   ' ADOコネクションを作成

    On Error GoTo ErrorTrap
    
    '192.168.11.11 自宅
    '10.206.3.155 学校
    '192.168.100.102 木村
    
    ' ODBC接続
    adoCon.Open _
        "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
        " SERVER=10.206.3.155;" & _
        " DATABASE=arisan_db;" & _
        " UID=test;" & _
        " PWD=Test123_pass;"
    
    ' クライアントサイドカーソルに変更
    adoCon.CursorLocation = 3
    
    ' 接続成功
    connect = True
Exit Function

ErrorTrap:
    ' 接続失敗
    connect = False

    ' 解放処理
    Set adoCon = Nothing
    
    ' エラー処理
    Select Case Err.Number
        ' DB接続エラー
        Case -2147467259
            MsgBox "データベースに接続できません"
    End Select
End Function

' データベースへの接続を切断
Public Sub disconnect()
    adoCon.Close
    Set adoCon = Nothing
End Sub

' SQL文実行
Public Function execute(SQL As String) As ADODB.Recordset
    Set adoRs = New ADODB.Recordset

    adoRs.Open SQL, adoCon
    
    ' 戻り値
    Set execute = adoRs
End Function
