Attribute VB_Name = "シフト表"
Option Explicit

Sub 希望休入力()
    ' フォームを表示する。
    シフト表Form.Show
End Sub

Sub 翌月のシート作成()
    ' シートから年月日を取得し変数に格納する。
    Dim d As Date       ' 日
    Dim y As Integer    ' 年
    Dim m As Integer    ' 月
    
    d = Range("I2").Value
    y = year(d)
    m = Month(d)
    

    ' 月を翌月に変更し、翌月のシート名を設定した後、翌月のシートが既に作成されているかを判定する。
    Dim NextSheet As String     ' 翌月のシート名
    Dim SameSheet As Boolean    ' 同名シート判定フラグ
        
    If m = 12 Then
        y = y + 1
        m = 1
    Else
        m = m + 1
    End If
    
    NextSheet = y & "." & m
    SameSheet = exist_check(NextSheet, "ws")
    
    
    ' 翌月のシートが存在する場合は、メッセージを表示する。
    ' 翌月のシートが存在しない場合は新規シートを作成し、シート名の変更・セルの内容の変更・セルのクリアを行う。
    If SameSheet = True Then
        MsgBox "翌月のシートは既に作成されています"
    Else
        Worksheets(ActiveSheet.name).Copy After:=ActiveSheet
        ActiveSheet.name = NextSheet
             
        Range("I2").Value = DateAdd("m", 1, d)
        
        Range(Cells(6, 9), Cells(56, 39)).ClearContents
    End If
End Sub

Sub シフト表コピー()
    ' コピー元ブック名と現在マクロを実行しているExcelブックからコピー先ブック名を設定する。
    ' コピー元ブック名は「卒論シフト仮生成.xlsm」、コピー先ブック名は「卒論システム.xlsm」である。
    Dim SourceBook As String ' コピー元ブック名
    Dim TargetBook As String ' コピー先ブック名
    
    SourceBook = "卒論シフト仮生成.xlsm"
    TargetBook = Application.ThisWorkbook.name
    
    
    ' コピー元ブックが開かれているか、コピー元ブックが読み取り専用であるかを確認し、条件を満たした場合メッセージを表示する。
    Dim SameBook As Boolean ' 同名ブック判定フラグ
    
    SameBook = exist_check(SourceBook, "wb")
    
    If SameBook = False Then
        MsgBox SourceBook & " が開かれていません" & vbCrLf & "ブックを開いてから実行してください"
        Exit Sub
    End If
    
    If Workbooks(SourceBook).ReadOnly = True Then
        MsgBox SourceBook & " が読み取り専用になっています" & vbCrLf & "編集可能にして実行してください"
        Exit Sub
    End If
    
    
    ' コピー元ブックとコピー先ブックでセルの範囲を指定し、コピー処理を行う。
    ' コピー処理後、コピーモードを解除する。
    Workbooks(SourceBook).Worksheets("勤務スケジュール修正").Range("E6:AI15").Copy
    Workbooks(TargetBook).Worksheets(ActiveSheet.name).Range("I6:AM15").PasteSpecial (xlPasteValues)
    
    Application.CutCopyMode = False
End Sub
