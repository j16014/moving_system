Attribute VB_Name = "�\���"
Option Explicit

' �ϐ�
Dim i As Integer
Dim j As Integer
Dim k As Integer

' �\��󋵎Q��
Sub Info_SELECT_click()
    Dim db As DBManager         ' �f�[�^�x�[�X�N���X
    Dim dbFlg As Boolean        ' �ڑ��t���O
    Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
    Dim SQL As String           ' SQL
    Dim FldCount As Integer     ' �t�B�[���h��
    Dim RecCount As Long        ' ���R�[�h��
    Dim myArray() As Variant    ' �Q�ƃ��R�[�h�z��
    Dim MonthSheet As String    ' �����Ƃ̃V�[�g��
    Dim SameSheet As String     ' �����V�[�g
    Dim wDay As Integer         ' �j��
    Dim Col As Integer          ' �Z���Ԓn�̗�(A��)
    Dim Row As Integer          ' �Z���Ԓn�̍s(1�s)
    Dim searchE As String       ' �����v�f
    Dim arrayEnd As Integer     ' �z��̒���
    Dim moveCount As Integer    ' �\�񌏐�
    Dim workerCount As Integer  ' �Ј���
    Dim freeCount As Integer    ' �]�T�Ј���
    Dim ws As Worksheet         ' ���[�N�V�[�g
    
    ' �e���̃Z���̓h��Ԃ����N���A
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
        
        ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
        Set db = New DBManager
        dbFlg = db.connect
        
        ' SQL��
        SQL = "SELECT DATE_FORMAT(move_day, '%Y-%m-%d') AS time, COUNT(*) AS count " & _
        "FROM customers WHERE DATE_FORMAT(move_day, '%Y') = '" & Range("N1").Value & "' GROUP BY time;"
    
        ' SQL�̎��s
        Set adoRs = db.execute(SQL)
        
        ' �t�B�[���h���ƃ��R�[�h�����擾
        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
        
        ' ���R�[�h�������ꍇ
        If adoRs.EOF Then
            MsgBox Range("N1").Value & "�N�̂��q�l�f�[�^�͂���܂���B"
        Else
            ' �񎟌��z����Ē�`
            ReDim myArray(FldCount - 1, RecCount - 1)
            ' ���R�[�h�Z�b�g�̓��e��ϐ��Ɋi�[
            myArray = adoRs.GetRows
            
            ' �z��̍Ō���̓Y����
            arrayEnd = UBound(myArray, 2)
            
            ' 12�����[�v
            For i = 1 To 12
                MonthSheet = Range("N1").Value & "." & i
                
                SameSheet = False
                
                ' �u�b�N���Ƀ��[�N�V�[�g�����邩����
                For Each ws In Sheets
                    If ws.name = MonthSheet Then
                        ' ���݂���
                        SameSheet = True
                    End If
                Next
                 
                If SameSheet = True Then
                    ' �����Ƃ̏��T�̔Ԓn�ݒ�
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
                    
                    ' ���̏����̗j�������
                    wDay = Weekday(DateSerial(Range("N1"), i, 1))
                    
                    ' 31�����[�v
                    For j = 1 To 31
                    
                        ' ���̏����̔Ԓn��ݒ�
                        If j = 1 Then
                            Col = Col + (wDay - 1)
                        End If
                        
                        ' ��������v�f���w��
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
                        
                        ' �z��Ɉ�v���镶���񂪂��邩����
                        For k = 0 To arrayEnd
                            If myArray(0, k) = searchE Then
                                ' �\�񌏐����擾
                                moveCount = CInt(myArray(1, k))
                                
                                ' �Ј������擾
                                workerCount = Worksheets(MonthSheet).Cells(58, 8 + j).Value
                                    
                                ' �]�T�Ј������v�Z
                                freeCount = workerCount * 2 - moveCount

                                ' �Z���̐F��ς���
                                If freeCount <= 0 Then
                                    ' �ԐF�œh��Ԃ�
                                    Cells(Row, Col).Interior.ColorIndex = 3
                                ElseIf freeCount <= 3 Then
                                    ' ���F�œh��Ԃ�
                                    Cells(Row, Col).Interior.ColorIndex = 6
                                End If
                            End If
                        Next k
                        
                        ' �y�j���ŉ��s�A����ȊO�͗j����������炷
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
        
        ' �������
        adoRs.Close
        Set adoRs = Nothing
        db.disconnect
        Set db = Nothing
    Else
        MsgBox "���͂����l�͐���������܂���"
    End If
Exit Sub
    
ErrorTrap:
    ' �������
    Set adoRs = Nothing
    Set db = Nothing
    
    ' �G���[����
    Select Case Err.Number
        ' DB�ڑ��G���[
        Case -2147467259
            MsgBox "�f�[�^�x�[�X�ɐڑ��ł��܂���"
    End Select
End Sub
