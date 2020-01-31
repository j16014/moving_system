Attribute VB_Name = "予約状況"
Option Explicit

' 変数
Dim i As Integer
Dim j As Integer
Dim k As Integer

' 予約状況参照
Sub Info_SELECT_click()
    ' 各月のセルの塗りつぶしをクリアする。
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


    ' N1セルが正当な値の場合は処理を実行する。
    If Range("N1").Value = "2019" Or _
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
    
        
        ' DBManagerクラスのdbをインスタンス化し、接続処理を行う
        Dim db As DBManager         ' データベースクラス
        
        On Error GoTo ErrorTrap
        
        Set db = New DBManager
        db.connect
        
        
        ' SQL文を実行し､レコードセットからフィールド数とレコード数を取得する｡
        Dim SQL As String           ' SQL
        Dim adoRs As Object         ' ADOレコードセット
        Dim FldCount As Integer     ' フィールド数
        Dim RecCount As Long        ' レコード数
        
        SQL = "SELECT DATE_FORMAT(move_day, '%Y-%m-%d') AS time, COUNT(*) AS count " & _
        "FROM customers WHERE DATE_FORMAT(move_day, '%Y') = '" & Range("N1").Value & "' GROUP BY time;"
    
        Set adoRs = db.execute(SQL)
        
        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
        
        
        ' レコードが存在しない場合はメッセージを表示する。
        ' レコードが存在する場合はレコードセットを配列に格納する。そして配列の最後尾の添え字を取得することでレコードの件数を取得する。
        Dim myArray() As Variant    ' 参照レコード配列
        Dim arrayEnd As Integer     ' 配列の長さ

        If adoRs.EOF Then
            MsgBox Range("N1").Value & "年のお客様データはありません。"
        Else
            ReDim myArray(FldCount - 1, RecCount - 1)
            myArray = adoRs.GetRows

            arrayEnd = UBound(myArray, 2)
            
            
            ' 12ヶ月分繰り返し処理を行う。シート名を設定して同名のシートが存在するかを判定する。
            Dim MonthSheet As String    ' 月ごとのシート名
            Dim SameSheet As String     ' 同名シート
        
            For i = 1 To 12
                MonthSheet = Range("N1").Value & "." & i
                
                SameSheet = exist_check(MonthSheet, "ws")
                
                
                ' 同名のシートが存在する場合、月毎の初週の番地を設定する。またWeekday関数より初日の曜日を取得する。
                Dim Col As Integer          ' セル番地の列(A列)
                Dim Row As Integer          ' セル番地の行(1行)
                Dim wDay As Integer         ' 曜日
                
                If SameSheet = True Then
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
                    
                    wDay = Weekday(DateSerial(Range("N1"), i, 1))
                    
                    
                    ' 31日分繰り返し処理を行う。月の初週の番地と初日の曜日から初日の番地を設定する。
                    For j = 1 To 31
                        If j = 1 Then
                            Col = Col + (wDay - 1)
                        End If
                        
                        
                        ' セルの値から検索文字列を設定する。
                        Dim searchE As String       ' 検索文字列
                        
                        If i < 10 And j < 10 Then
                            searchE = Range("N1").Value & "-0" & i & "-0" & j
                        ElseIf i < 10 Then
                            searchE = Range("N1").Value & "-0" & i & "-" & j
                        ElseIf j < 10 Then
                            searchE = Range("N1").Value & "-" & i & "-0" & j
                        Else
                            searchE = Range("N1").Value & "-" & i & "-" & j
                        End If
                        
                        
                        ' 配列の長さ分繰り返し処理を行う。検索文字列と一致する要素が配列に存在する場合、配列から予約件数、シフト表から社員数を取得し、残り社員数を計算する。
                        Dim moveCount As Integer    ' 予約件数
                        Dim workerCount As Integer  ' 社員数
                        Dim freeCount As Integer    ' 残り社員数
                        
                        moveCount = 0
                        
                        For k = 0 To arrayEnd
                            If myArray(0, k) = searchE Then
                                moveCount = CInt(myArray(1, k))
                                
                                workerCount = Worksheets(MonthSheet).Cells(58, 8 + j).Value
                                
                                freeCount = workerCount * 2 - moveCount


                                ' 残り社員数が0以下の場合赤、3以下の場合黄にセルの塗りつぶしを設定する。
                                If freeCount <= 0 Then
                                    Cells(Row, Col).Interior.ColorIndex = 3
                                ElseIf freeCount <= 3 Then
                                    Cells(Row, Col).Interior.ColorIndex = 6
                                End If
                            End If
                        Next k
                        

                        ' 日付を一日進めると同時に曜日を一日ずらす必要があるため、セルの番地変数を設定する。
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
        
        
        ' レコードセットとデータベースオブジェクトの解放処理を行う。
        adoRs.Close
        Set adoRs = Nothing
        db.disconnect
        Set db = Nothing
        
        
    ' N1セルが不正な値の場合はメッセージを表示する。
    Else
        MsgBox "入力した値は正しくありません。"
    End If
Exit Sub
    
ErrorTrap:
    ' エラーが発生した場合、レコードセットとデータベースオブジェクトの解放処理を行う。
    Set adoRs = Nothing
    Set db = Nothing
End Sub
