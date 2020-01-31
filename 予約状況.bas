Attribute VB_Name = "�\���"
Option Explicit

' �ϐ�
Dim i As Integer
Dim j As Integer
Dim k As Integer

' �\��󋵎Q��
Sub Info_SELECT_click()
    ' �e���̃Z���̓h��Ԃ����N���A����B
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


    ' N1�Z���������Ȓl�̏ꍇ�͏��������s����B
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
    
        
        ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
        Dim db As DBManager         ' �f�[�^�x�[�X�N���X
        
        On Error GoTo ErrorTrap
        
        Set db = New DBManager
        db.connect
        
        
        ' SQL�������s������R�[�h�Z�b�g����t�B�[���h���ƃ��R�[�h�����擾����
        Dim SQL As String           ' SQL
        Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
        Dim FldCount As Integer     ' �t�B�[���h��
        Dim RecCount As Long        ' ���R�[�h��
        
        SQL = "SELECT DATE_FORMAT(move_day, '%Y-%m-%d') AS time, COUNT(*) AS count " & _
        "FROM customers WHERE DATE_FORMAT(move_day, '%Y') = '" & Range("N1").Value & "' GROUP BY time;"
    
        Set adoRs = db.execute(SQL)
        
        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
        
        
        ' ���R�[�h�����݂��Ȃ��ꍇ�̓��b�Z�[�W��\������B
        ' ���R�[�h�����݂���ꍇ�̓��R�[�h�Z�b�g��z��Ɋi�[����B�����Ĕz��̍Ō���̓Y�������擾���邱�ƂŃ��R�[�h�̌������擾����B
        Dim myArray() As Variant    ' �Q�ƃ��R�[�h�z��
        Dim arrayEnd As Integer     ' �z��̒���

        If adoRs.EOF Then
            MsgBox Range("N1").Value & "�N�̂��q�l�f�[�^�͂���܂���B"
        Else
            ReDim myArray(FldCount - 1, RecCount - 1)
            myArray = adoRs.GetRows

            arrayEnd = UBound(myArray, 2)
            
            
            ' 12�������J��Ԃ��������s���B�V�[�g����ݒ肵�ē����̃V�[�g�����݂��邩�𔻒肷��B
            Dim MonthSheet As String    ' �����Ƃ̃V�[�g��
            Dim SameSheet As String     ' �����V�[�g
        
            For i = 1 To 12
                MonthSheet = Range("N1").Value & "." & i
                
                SameSheet = exist_check(MonthSheet, "ws")
                
                
                ' �����̃V�[�g�����݂���ꍇ�A�����̏��T�̔Ԓn��ݒ肷��B�܂�Weekday�֐���菉���̗j�����擾����B
                Dim Col As Integer          ' �Z���Ԓn�̗�(A��)
                Dim Row As Integer          ' �Z���Ԓn�̍s(1�s)
                Dim wDay As Integer         ' �j��
                
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
                    
                    
                    ' 31�����J��Ԃ��������s���B���̏��T�̔Ԓn�Ə����̗j�����珉���̔Ԓn��ݒ肷��B
                    For j = 1 To 31
                        If j = 1 Then
                            Col = Col + (wDay - 1)
                        End If
                        
                        
                        ' �Z���̒l���猟���������ݒ肷��B
                        Dim searchE As String       ' ����������
                        
                        If i < 10 And j < 10 Then
                            searchE = Range("N1").Value & "-0" & i & "-0" & j
                        ElseIf i < 10 Then
                            searchE = Range("N1").Value & "-0" & i & "-" & j
                        ElseIf j < 10 Then
                            searchE = Range("N1").Value & "-" & i & "-0" & j
                        Else
                            searchE = Range("N1").Value & "-" & i & "-" & j
                        End If
                        
                        
                        ' �z��̒������J��Ԃ��������s���B����������ƈ�v����v�f���z��ɑ��݂���ꍇ�A�z�񂩂�\�񌏐��A�V�t�g�\����Ј������擾���A�c��Ј������v�Z����B
                        Dim moveCount As Integer    ' �\�񌏐�
                        Dim workerCount As Integer  ' �Ј���
                        Dim freeCount As Integer    ' �c��Ј���
                        
                        moveCount = 0
                        
                        For k = 0 To arrayEnd
                            If myArray(0, k) = searchE Then
                                moveCount = CInt(myArray(1, k))
                                
                                workerCount = Worksheets(MonthSheet).Cells(58, 8 + j).Value
                                
                                freeCount = workerCount * 2 - moveCount


                                ' �c��Ј�����0�ȉ��̏ꍇ�ԁA3�ȉ��̏ꍇ���ɃZ���̓h��Ԃ���ݒ肷��B
                                If freeCount <= 0 Then
                                    Cells(Row, Col).Interior.ColorIndex = 3
                                ElseIf freeCount <= 3 Then
                                    Cells(Row, Col).Interior.ColorIndex = 6
                                End If
                            End If
                        Next k
                        

                        ' ���t������i�߂�Ɠ����ɗj����������炷�K�v�����邽�߁A�Z���̔Ԓn�ϐ���ݒ肷��B
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
        
        
        ' ���R�[�h�Z�b�g�ƃf�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
        adoRs.Close
        Set adoRs = Nothing
        db.disconnect
        Set db = Nothing
        
        
    ' N1�Z�����s���Ȓl�̏ꍇ�̓��b�Z�[�W��\������B
    Else
        MsgBox "���͂����l�͐���������܂���B"
    End If
Exit Sub
    
ErrorTrap:
    ' �G���[�����������ꍇ�A���R�[�h�Z�b�g�ƃf�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
    Set adoRs = Nothing
    Set db = Nothing
End Sub
