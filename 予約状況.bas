Attribute VB_Name = "予約状況"
Option Explicit

' 変数
Dim i As Integer
Dim j As Integer
Dim k As Integer

' 予約状況参照
Sub Info_SELECT_click()
    Dim db As DBManager         ' データベースクラス
    Dim dbFlg As Boolean        ' 接続フラグ
    Dim adoRs As Object         ' ADOレコードセット
    Dim SQL As String           ' SQL
    Dim FldCount As Integer     ' フィールド数
    Dim RecCount As Long        ' レコード数
    Dim myArray() As Variant    ' 参照レコード配列
    Dim MonthSheet As String    ' 月ごとのシート名
    Dim SameSheet As String     ' 同名シート
    Dim wDay As Integer         ' 曜日
    Dim Col As Integer          ' セル番地の列(A列)
    Dim Row As Integer          ' セル番地の行(1行)
    Dim searchE As String       ' 検索要素
    Dim arrayEnd As Integer     ' 配列の長さ
    Dim moveCount As Integer    ' 予約件数
    Dim workerCount As Integer  ' 社員数
    Dim freeCount As Integer    ' 余裕社員数
    Dim ws As Worksheet         ' ワークシート
    
    ' 各月のセルの塗りつぶしをクリア
    Range("B5:H10").Interior.ColorIndex = 0
    Range("J5:P10").Interior.ColorIndex = 0
    Range("R5:X10").Interior.ColorIndex = 0
    Range("Z5:AF10").Interior.ColorIndex = 0
    Range("B14:H19").Interior.ColorIndex = 0
    Range("J14:P19").Interior.ColorIndex = 0
    Range("R14:X19").Interior.ColorIndex = 0
    Range("Z14:AF19").Interior.ColorIndex = 0
    Range("B23:H28").Interior.ColorIndex = 0
    Range("J23:P28").Interior.ColorIndex = 0
    Range("R23:X28").Interior.ColorIndex = 0
    Range("Z23:AF28").Interior.ColorIndex = 0

    If _
    Range("N1").Value = "2019" Or _
    Range("N1").Value = "2020" Or _
    Range("N1").Value = "2021" Or _
    Range("N1").Value = "2022" Or _
    Range("N1").Value = "2023" Or _
    Range("N1").Value = "2024" Or _
    Range("N1").Value = "2025" Or _
    Range("N1").Value = "2026" Or _
    Range("N1").Value = "2027" Or _
    Range("N1").Value = "2028" Or _
    Range("N1").Value = "2029" Or _
    Range("N1").Value = "2030" Then
    
        On Error GoTo ErrorTrap
        
        ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
        Set db = New DBManager
        dbFlg = db.connect
        
        ' SQL文
        SQL = "SELECT DATE_FORMAT(move_day, '%Y-%m-%d') AS time, COUNT(*) AS count " & _
        "FROM customers WHERE DATE_FORMAT(move_day, '%Y') = '" & Range("N1").Value & "' GROUP BY time;"
    
        ' SQLの実行
        Set adoRs = db.execute(SQL)
        
        ' フィールド数とレコード数を取得
        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
        
        ' レコードが無い場合
        If adoRs.EOF Then
            MsgBox Range("N1").Value & "年のお客様データはありません。"
        Else
            ' 二次元配列を再定義
            ReDim myArray(FldCount - 1, RecCount - 1)
            ' レコードセットの内容を変数に格納
            myArray = adoRs.GetRows
            
            ' 配列の最後尾の添え字
            arrayEnd = UBound(myArray, 2)
            
            ' 12月ループ
            For i = 1 To 12
                MonthSheet = Range("N1").Value & "." & i
                
                SameSheet = False
                
                ' ブック内にワークシートがあるか検索
                For Each ws In Sheets
                    If ws.name = MonthSheet Then
                        ' 存在する
                        SameSheet = True
                    End If
                Next
                 
                If SameSheet = True Then
                    ' 月ごとの初週の番地設定
                    If i <= 4 Then
                        Col = 2 + 8 * (i - 1)
                        Row = 5
                    ElseIf i >= 5 And i <= 8 Then
                        Col = 2 + 8 * (i - 5)
                        Row = 14
                    ElseIf i >= 9 Then
                        Col = 2 + 8 * (i - 9)
                        Row = 23
                    End If
                    
                    ' 月の初日の曜日を特定
                    wDay = Weekday(DateSerial(Range("N1"), i, 1))
                    
                    ' 31日ループ
                    For j = 1 To 31
                    
                        ' 月の初日の番地を設定
                        If j = 1 Then
                            Col = Col + (wDay - 1)
                        End If
                        
                        ' 検索する要素を指定
                        If i < 10 And j < 10 Then
                            searchE = Range("N1").Value & "-0" & i & "-0" & j
                        ElseIf i < 10 Then
                            searchE = Range("N1").Value & "-0" & i & "-" & j
                        ElseIf j < 10 Then
                            searchE = Range("N1").Value & "-" & i & "-0" & j
                        Else
                            searchE = Range("N1").Value & "-" & i & "-" & j
                        End If
                        
                        moveCount = 0
                        
                        ' 配列に一致する文字列があるか検索
                        For k = 0 To arrayEnd
                            If myArray(0, k) = searchE Then
                                ' 予約件数を取得
                                moveCount = CInt(myArray(1, k))
                                
                                ' 社員数を取得
                                workerCount = Worksheets(MonthSheet).Cells(58, 8 + j).Value
                                    
                                ' 余裕社員数を計算
                                freeCount = workerCount * 2 - moveCount

                                ' セルの色を変える
                                If freeCount <= 0 Then
                                    ' 赤色で塗りつぶす
                                    Cells(Row, Col).Interior.ColorIndex = 3
                                ElseIf freeCount <= 3 Then
                                    ' 黄色で塗りつぶす
                                    Cells(Row, Col).Interior.ColorIndex = 6
                                End If
                            End If
                        Next k
                        
                        ' 土曜日で改行、それ以外は曜日を一日ずらす
                        If wDay = 7 Then
                            wDay = 1
                            Col = Col - 6
                            Row = Row + 1
                        Else
                            Col = Col + 1
                            wDay = wDay + 1
                        End If
                    Next j
                End If
            Next i
        End If
        
        ' 解放処理
        adoRs.Close
        Set adoRs = Nothing
        db.disconnect
        Set db = Nothing
    Else
        MsgBox "入力した値は正しくありません"
    End If
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
