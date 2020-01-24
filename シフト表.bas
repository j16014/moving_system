Attribute VB_Name = "シフト表"
Option Explicit

Sub 希望休入力()
    シフト表Form.Show
End Sub

Sub 翌月のシート作成()
    Dim d As Date               ' シートの日時
    Dim y As Integer            ' 年変数
    Dim m As Integer            ' 月変数
    Dim NextSheet As String     ' 作成したシート名
    Dim SameSheet As Boolean    ' 同名シート判定フラグ
    Dim ws As Worksheet         ' ワークシート
    Dim SheetName As String     ' ワークシート名
    Dim NextMonth As Date       ' 翌月
    
    ' I2の年月日を取得
    d = Range("I2").Value
    y = year(d)
    m = Month(d)
    
    ' 12月なら年度を変える
    If m = 12 Then
        y = y + 1
        m = 1
    Else
        m = m + 1
    End If
    
    ' 翌月のシート名作成
    NextSheet = y & "." & m
    
    ' 翌月のシートが存在したらメッセージ、存在しなければ翌月のシート作成
    SameSheet = False
    
    ' ブック内のワークシート名を取得
    For Each ws In Sheets
        If ws.name = NextSheet Then
            ' 存在する
            SameSheet = True
        End If
    Next
    
    If SameSheet = True Then
        MsgBox "翌月のシートは既に作成されています"
        Exit Sub
    Else
        ' ワークシート名を取得
        SheetName = ActiveSheet.name
        
        ' 翌月のシート作成
        Worksheets(SheetName).Copy After:=ActiveSheet
        ActiveSheet.name = NextSheet
             
        ' 月を一つ足す
        NextMonth = DateAdd("m", 1, d)
        Range("I2").Value = NextMonth
        
        ' セルの内容をクリアする
        Range(Cells(6, 9), Cells(56, 39)).ClearContents
    End If
End Sub

Sub シフト表コピー()
    
    Dim SourceFile As String    ' コピー元ファイル名
    Dim TargetFile As String    ' コピー先ファイル名
    Dim SameFile As Boolean     ' 同名ファイル判定フラグ
    Dim wb As Workbook          ' ワークブック
    
    ' コピー元ファイル名
    SourceFile = "卒論シフト仮生成.xlsm"
    ' コピー先ファイル名
    TargetFile = Application.ThisWorkbook.name
    
    ' ファイルが既に開いているかチェック
    SameFile = False
    
    For Each wb In Workbooks
        If wb.name = SourceFile Then
            ' 開いている
            SameFile = True
        End If
    Next wb
    
    ' コピー元ファイルが開いているかチェック
    If SameFile = False Then
        MsgBox SourceFile & " が開かれていません" & vbCrLf & "ファイルを開いてから実行してください"
        Exit Sub
    End If
    
    ' 読み取り専用チェック
    If Workbooks(SourceFile).ReadOnly = True Then
        MsgBox SourceFile & " が読み取り専用になっています" & vbCrLf & "編集可能にして実行してください"
        Exit Sub
    End If
    
    ' コピー処理
    Workbooks(SourceFile).Worksheets("勤務スケジュール修正").Range("E6:AI15").Copy
    Workbooks(TargetFile).Worksheets(ActiveSheet.name).Range("I6:AM15").PasteSpecial (xlPasteValues)
    
    ' コピーモード解除
    Application.CutCopyMode = False
End Sub
