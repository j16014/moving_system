Attribute VB_Name = "�z�ԕ\"
Option Explicit

' �ϐ�
Dim i As Integer
Dim j As Integer
Dim k As Integer

' �\��Q��
Sub Re_SELECT_Click()
    ' �Z���̓��e���N���A����B
    Range("C4:L55").Value = ""
    Range("P4:Y55").Value = ""


    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s���B
    Dim db As DBManager     ' �f�[�^�x�[�X�N���X
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    
    
    ' ���R�[�h�����݂��Ȃ������ꍇ�ɉ��Z���Ă����t���O���`����B
    Dim Eofnum As Integer   ' �f�[�^��������t���O
    Eofnum = 0
    

    ' ���z���ԑю�ʂ̎�ʐ����������J��Ԃ��A���̓s�xmeridian�ɈقȂ��ʂ�ݒ肷��B
    Dim meridian As String  ' ���z���ԑю�ʁiam,pm,free�j

    For i = 1 To 3
        Select Case i
            Case 1
                meridian = "AM"
            Case 2
                meridian = "PM"
            Case 3
                meridian = "free"
        End Select
        
        
        ' SQL�������s���ă��R�[�h�Z�b�g����t�B�[���h���ƃ��R�[�h�����擾����B
        Dim SQL As String           ' SQL
        Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
        Dim FldCount As Integer     ' �t�B�[���h��
        Dim RecCount As Long        ' ���R�[�h��

        SQL = "SELECT id,name,meridian,now_address,now_floors,now_ev,now_type," & _
        "new_address,new_floors,new_ev,new_type,preview_name,point," & _
        "start_time1,start_time2,start_time3,plan,difficulty,truck,driver," & _
        "assistant1,assistant2,assistant3,assistant4 FROM customers WHERE " & _
        "move_day = '" & Range("J1").Value & "-" & Range("M1").Value & "-" & Range("Q1").Value & "' " & _
        "AND meridian = '" & meridian & "'"
        
        Set adoRs = db.execute(SQL)

        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
            

        ' ���R�[�h�����݂��Ȃ��ꍇ�̓t���O�����Z����B
        ' ���R�[�h�����݂���ꍇ�́A���R�[�h�Z�b�g��z��Ɋi�[����B
        Dim myArray() As Variant    ' �Q�ƃ��R�[�h�z��
        
        If adoRs.EOF Then
            Select Case i
                Case 1
                    Eofnum = Eofnum + 1
                Case 2
                    Eofnum = Eofnum + 1
                Case 3
                    Eofnum = Eofnum + 1
            End Select
        Else
            ReDim myArray(FldCount - 1, RecCount - 1)
            myArray = adoRs.GetRows
                           
                    
            ' ���R�[�h�̌������E�Z���̕������ݒ肷�邽�߁A���R�[�h�����������J��Ԃ��B
            Dim condition As String     ' ���E�V�������
            
            For j = 0 To RecCount - 1
                condition = ""
                            
                   
                ' �e�ϐ��ɔz�񂩂�Y������l���擾���A�ו���ςޒn�_�̌�������ݒ肷��B
                Dim now_floors As String    ' ���K�w
                Dim now_ev As String        ' ��ev
                Dim now_type As String      ' ���������
                Dim flg As Boolean          ' �G���x�[�^�t���O
                            
                now_floors = myArray(4, j)
                now_ev = myArray(5, j)
                now_type = myArray(6, j)
                flg = False
                           
                Select Case now_type
                    Case "�A�p�[�g", "�c�n", "MC"
                        condition = condition & now_floors
                        flg = True
                    Case "�Б�", "�ꌬ��"
                        Select Case now_floors
                           Case 1
                                condition = condition & "1"
                           Case 2
                                condition = condition & "1/2"
                           Case 3
                                condition = condition & "1/2/3"
                           Case 4
                                condition = condition & "1/2/3/4"
                           Case Else
                                condition = condition & now_floors
                    End Select
                End Select
                        
                If flg = True Then
                    If now_ev = "EV�L" Then
                        condition = condition & "���`"
                    Else
                        condition = condition & "�~�`"
                    End If
                Else
                    condition = condition & "�`"
                End If
                    
                    
                ' �e�ϐ��ɔz�񂩂�Y������l���擾���A�ו����~�낷�n�_�̌�������ݒ肷��B
                Dim new_floors As String    ' �V�K�w
                Dim new_ev As String        ' �Vev
                Dim new_type As String      ' �V�������
                
                new_floors = myArray(8, j)
                new_ev = myArray(9, j)
                new_type = myArray(10, j)
                flg = False
               
                Select Case new_type
                    Case "�A�p�[�g", "�c�n", "MC"
                        condition = condition & new_floors
                        flg = True
                    Case "���V�z", "�Б�", "�ꌬ��"
                        Select Case new_floors
                            Case 1
                                condition = condition & "1"
                            Case 2
                                condition = condition & "1/2"
                            Case 3
                                condition = condition & "1/2/3"
                            Case 4
                                condition = condition & "1/2/3/4"
                            Case Else
                                condition = condition & new_floors
                        End Select
                End Select
                        
                If flg = True Then
                    If new_ev = "EV�L" Then
                        condition = condition & "��"
                    Else
                        condition = condition & "�~"
                    End If
                End If
    
                
                ' �e�ϐ��ɔz�񂩂�Y������l���擾���A�Z����ݒ肷��B
                Dim address As String       ' ���E�V�Z��
                Dim now_address As String   ' ���Z��
                Dim new_address As String   ' �V�Z��
                
                address = ""
                        
                now_address = myArray(3, j)
                new_address = myArray(7, j)
                address = now_address & " �` " & new_address
                
                
                ' �z��̓��e���Z���Ɋi�[����B���z���ԑю�ʂɂ���ăZ���̔Ԓn���ς�邽�ߏ������򂵂Ă���B
                If i = 1 Then
                    Range("E" & j * 4 + 4) = myArray(0, j)     ' ID
                    Range("G" & j * 4 + 4) = myArray(1, j)     ' ���q�l����
                    Range("E" & j * 4 + 6) = condition         ' ���V�������
                    Range("F" & j * 4 + 6) = address           ' ���V�Z��
                    Range("D" & j * 4 + 6) = myArray(11, j)    ' �����S��
                    Range("F" & j * 4 + 4) = myArray(12, j)    ' �|�C���g��
                    Range("C" & j * 4 + 4) = myArray(13, j)    ' �J�n����1
                    Range("C" & j * 4 + 5) = myArray(14, j)    ' �J�n����2
                    Range("C" & j * 4 + 7) = myArray(15, j)    ' �J�n����3
                    Range("D" & j * 4 + 4) = myArray(16, j)    ' �v����
                    Range("I" & j * 4 + 4) = myArray(17, j)    ' ��Փx
                    Range("J" & j * 4 + 4) = myArray(18, j)    ' �g���b�N
                    Range("J" & j * 4 + 6) = myArray(19, j)    ' �h���C�o�[
                    Range("K" & j * 4 + 4) = myArray(20, j)    ' ����1
                    Range("K" & j * 4 + 6) = myArray(21, j)    ' ����2
                    Range("L" & j * 4 + 4) = myArray(22, j)    ' ����3
                    Range("L" & j * 4 + 6) = myArray(23, j)    ' ����4
                ElseIf i = 2 Then
                    Range("R" & j * 4 + 4) = myArray(0, j)     ' ID
                    Range("T" & j * 4 + 4) = myArray(1, j)     ' ���q�l����
                    Range("R" & j * 4 + 6) = condition         ' ���V�������
                    Range("S" & j * 4 + 6) = address           ' ���V�Z��
                    Range("Q" & j * 4 + 6) = myArray(11, j)    ' �����S��
                    Range("S" & j * 4 + 4) = myArray(12, j)    ' �|�C���g��
                    Range("P" & j * 4 + 4) = myArray(13, j)    ' �J�n����1
                    Range("P" & j * 4 + 5) = myArray(14, j)    ' �J�n����2
                    Range("P" & j * 4 + 7) = myArray(15, j)    ' �J�n����3
                    Range("Q" & j * 4 + 4) = myArray(16, j)    ' �v����
                    Range("V" & j * 4 + 4) = myArray(17, j)    ' ��Փx
                    Range("W" & j * 4 + 4) = myArray(18, j)    ' �g���b�N
                    Range("W" & j * 4 + 6) = myArray(19, j)    ' �h���C�o�[
                    Range("X" & j * 4 + 4) = myArray(20, j)    ' ����1
                    Range("X" & j * 4 + 6) = myArray(21, j)    ' ����2
                    Range("Y" & j * 4 + 4) = myArray(22, j)    ' ����3
                    Range("Y" & j * 4 + 6) = myArray(23, j)    ' ����4
                ElseIf i = 3 Then
                    If j < 5 Then
                        Range("E" & j * 4 + 36) = myArray(0, j)     ' ID
                        Range("G" & j * 4 + 36) = myArray(1, j)     ' ���q�l����
                        Range("E" & j * 4 + 38) = condition         ' ���V�������
                        Range("F" & j * 4 + 38) = address           ' ���V�Z��
                        Range("D" & j * 4 + 38) = myArray(11, j)    ' �����S��
                        Range("F" & j * 4 + 36) = myArray(12, j)    ' �|�C���g��
                        Range("C" & j * 4 + 36) = myArray(13, j)    ' �J�n����1
                        Range("C" & j * 4 + 37) = myArray(14, j)    ' �J�n����2
                        Range("C" & j * 4 + 39) = myArray(15, j)    ' �J�n����3
                        Range("D" & j * 4 + 36) = myArray(16, j)    ' �v����
                        Range("I" & j * 4 + 36) = myArray(17, j)    ' ��Փx
                        Range("J" & j * 4 + 36) = myArray(18, j)    ' �g���b�N
                        Range("J" & j * 4 + 38) = myArray(19, j)    ' �h���C�o�[
                        Range("K" & j * 4 + 36) = myArray(20, j)    ' ����1
                        Range("K" & j * 4 + 38) = myArray(21, j)    ' ����2
                        Range("L" & j * 4 + 36) = myArray(22, j)    ' ����3
                        Range("L" & j * 4 + 38) = myArray(23, j)    ' ����4
                    Else
                        Range("R" & j * 4 + 16) = myArray(0, j)     ' ID
                        Range("T" & j * 4 + 16) = myArray(1, j)     ' ���q�l����
                        Range("R" & j * 4 + 18) = condition         ' ���V�������
                        Range("S" & j * 4 + 18) = address           ' ���V�Z��
                        Range("Q" & j * 4 + 18) = myArray(11, j)    ' �����S��
                        Range("S" & j * 4 + 16) = myArray(12, j)    ' �|�C���g��
                        Range("P" & j * 4 + 16) = myArray(13, j)    ' �J�n����1
                        Range("P" & j * 4 + 17) = myArray(14, j)    ' �J�n����2
                        Range("P" & j * 4 + 19) = myArray(15, j)    ' �J�n����3
                        Range("Q" & j * 4 + 16) = myArray(16, j)    ' �v����
                        Range("V" & j * 4 + 16) = myArray(17, j)    ' ��Փx
                        Range("W" & j * 4 + 16) = myArray(18, j)    ' �g���b�N
                        Range("W" & j * 4 + 18) = myArray(19, j)    ' �h���C�o�[
                        Range("X" & j * 4 + 16) = myArray(20, j)    ' ����1
                        Range("X" & j * 4 + 18) = myArray(21, j)    ' ����2
                        Range("Y" & j * 4 + 16) = myArray(22, j)    ' ����3
                        Range("Y" & j * 4 + 18) = myArray(23, j)    ' ����4
                    End If
                End If
            Next j
        End If
    Next i
    
    
    ' ���R�[�h�����݂��Ȃ��ꍇ�̓��b�Z�[�W�\������B
    If Eofnum = 3 Then
        MsgBox "���q�l�f�[�^������܂���"
    End If
    
    
    ' ���R�[�h�Z�b�g�ƃf�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
    adoRs.Close
    Set adoRs = Nothing
    db.disconnect
    Set db = Nothing
Exit Sub
 
ErrorTrap:
    ' �G���[�����������ꍇ�A���R�[�h�Z�b�g�ƃf�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
    Set adoRs = Nothing
    Set db = Nothing
End Sub

' �\��X�V
Sub Re_UPDATA_Click()
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Dim db As DBManager     ' �f�[�^�x�[�X�N���X
    
    On Error GoTo ErrorTrap
   
    Set db = New DBManager
    db.connect


    ' �m�F�{�^����\����Yes�̏ꍇ�́A�z�ԕ\�V�[�g�̃Z���̒l��SQL����ݒ肵UPDATE�����s����B�V�[�g�͓��ɂȂ��Ă��邽�ߏ������򂷂邱�ƂŔԒn��ς��Ă���B
    Dim result As Long      ' YesNo�{�^���\��
    Dim SQL As String       ' SQL

    result = MsgBox("�㏑���ۑ����Ă���낵���ł����H", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        For i = 0 To 1
            For j = 0 To 13
                If i = 0 Then
                    If Range("E" & j * 4 + 4).Value <> "" Then
                        SQL = "UPDATE customers SET start_time1 = '" & Range("C" & j * 4 + 4).Value & "'," & _
                        "start_time2 = '" & Range("C" & j * 4 + 5).Value & "'," & _
                        "start_time3 = '" & Range("C" & j * 4 + 7).Value & "'," & _
                        "plan = '" & Range("D" & j * 4 + 4).Value & "'," & _
                        "difficulty = '" & Range("I" & j * 4 + 4).Value & "'," & _
                        "truck = '" & Range("J" & j * 4 + 4).Value & "'," & _
                        "driver = '" & Range("J" & j * 4 + 6).Value & "'," & _
                        "assistant1 = '" & Range("K" & j * 4 + 4).Value & "'," & _
                        "assistant2 = '" & Range("K" & j * 4 + 6).Value & "'," & _
                        "assistant3 = '" & Range("L" & j * 4 + 4).Value & "'," & _
                        "assistant4 = '" & Range("L" & j * 4 + 6).Value & "'" & _
                        " WHERE id = '" & Range("E" & j * 4 + 4).Value & "'"

                        db.execute SQL
                    End If
                Else
                    If Range("R" & j * 4 + 4).Value <> "" Then
                        SQL = "UPDATE customers SET start_time1 = '" & Range("P" & j * 4 + 4).Value & "'," & _
                        "start_time2 = '" & Range("P" & j * 4 + 5).Value & "'," & _
                        "start_time3 = '" & Range("P" & j * 4 + 7).Value & "'," & _
                        "plan = '" & Range("Q" & j * 4 + 4).Value & "'," & _
                        "difficulty = '" & Range("V" & j * 4 + 4).Value & "'," & _
                        "truck = '" & Range("W" & j * 4 + 4).Value & "'," & _
                        "driver = '" & Range("W" & j * 4 + 6).Value & "'," & _
                        "assistant1 = '" & Range("X" & j * 4 + 4).Value & "'," & _
                        "assistant2 = '" & Range("X" & j * 4 + 6).Value & "'," & _
                        "assistant3 = '" & Range("Y" & j * 4 + 4).Value & "'," & _
                        "assistant4 = '" & Range("Y" & j * 4 + 6).Value & "'" & _
                        " WHERE id = '" & Range("R" & j * 4 + 4).Value & "'"
                        
                        db.execute SQL
                    End If
                End If
            Next j
        Next i
    End If
 
 
    ' �f�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
    db.disconnect
    Set db = Nothing
Exit Sub
 
ErrorTrap:
    ' �G���[�����������ꍇ�A�f�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
    Set db = Nothing
End Sub
