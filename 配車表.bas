Attribute VB_Name = "�z�ԕ\"
Option Explicit

' �ϐ�
Dim i As Integer
Dim j As Integer
Dim k As Integer

' �\��Q��
Sub Re_SELECT_Click()
    Dim db As DBManager         ' �f�[�^�x�[�X�N���X
    Dim dbFlg As Boolean        ' �ڑ��t���O
    Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
    Dim SQL As String           ' SQL
    Dim FldCount As Integer     ' �t�B�[���h��
    Dim RecCount As Long        ' ���R�[�h��
    Dim myArray() As Variant    ' �Q�ƃ��R�[�h�z��
    Dim condition As String     ' ���V�������
    Dim address As String       ' ���V�Z��
    Dim flg As Boolean          ' �����Eev�t���O
    Dim Eofnum As Integer       ' �f�[�^��������
    Dim now_address As String   ' ���Z��
    Dim now_floors As String    ' ���K�w
    Dim now_ev As String        ' ��ev
    Dim now_type As String      ' ���������
    Dim new_address As String   ' �V�Z��
    Dim new_floors As String    ' �V�K�w
    Dim new_ev As String        ' �Vev
    Dim new_type As String      ' �V�������
    Dim meridian As String      ' am,pm,free
     
    On Error GoTo ErrorTrap
   
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Set db = New DBManager
    dbFlg = db.connect
    
    ' �z�ԕ\�N���A
    Range("C4:L55").Value = ""
    Range("P4:Y55").Value = ""
    
    ' �f�[�^�������菉����
    Eofnum = 0
    
    ' SQL��
    SQL = "SELECT id,name,meridian,now_address,now_floors,now_ev,now_type," & _
    "new_address,new_floors,new_ev,new_type,preview_name,point,start_time1,start_time2,start_time3," & _
    "plan,difficulty,truck,driver,assistant1,assistant2,assistant3,assistant4 FROM customers " & _
    "WHERE move_day = '" & Range("J1").Value & "-" & Range("M1").Value & "-" & Range("Q1").Value & "'"
    
    ' SQL�̎��s
    Set adoRs = db.execute(SQL)
        
    For i = 1 To 3
        ' AM�EPM�Efree�ŏ�������
        Select Case i
            Case 1
                meridian = "AM"
            Case 2
                meridian = "PM"
            Case 3
                meridian = "free"
        End Select
        
        ' SQL��
        SQL = "SELECT id,name,meridian,now_address,now_floors,now_ev,now_type," & _
        "new_address,new_floors,new_ev,new_type,preview_name,point," & _
        "start_time1,start_time2,start_time3,plan,difficulty,truck,driver," & _
        "assistant1,assistant2,assistant3,assistant4 FROM customers WHERE " & _
        "move_day = '" & Range("J1").Value & "-" & Range("M1").Value & "-" & Range("Q1").Value & "' " & _
        "AND meridian = '" & meridian & "'"
        
        ' SQL�̎��s
        Set adoRs = db.execute(SQL)
            
        ' �t�B�[���h���ƃ��R�[�h�����擾
        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
            
        ' AM�EPM�Efree�Ƀ��R�[�h�������ꍇ���Z
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
            ' �񎟌��z����Ē�`
            ReDim myArray(FldCount - 1, RecCount - 1)
            ' ���R�[�h�Z�b�g�̓��e��ϐ��Ɋi�[
            myArray = adoRs.GetRows
                    
            For j = 0 To RecCount - 1
                ' ����������
                condition = ""
                flg = False
                            
                ' �����ϐ���`
                ' ���K�w
                now_floors = myArray(4, j)
                ' ��ev
                now_ev = myArray(5, j)
                ' ���������
                now_type = myArray(6, j)
                ' �V�K�w
                new_floors = myArray(8, j)
                ' �Vev
                new_ev = myArray(9, j)
                ' �V�������
                new_type = myArray(10, j)
                           
                ' �ςݒn�̏���
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
                        
                ' �ςݒnev�Ɓ`
                If flg = True Then
                    If now_ev = "EV�L" Then
                        condition = condition & "���`"
                    Else
                        condition = condition & "�~�`"
                    End If
                Else
                    condition = condition & "�`"
                End If
                            
                ' �t���O���Z�b�g
                flg = False
                        
                ' �~�낵�n����
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
                        
                ' �~�낵�nev
                If flg = True Then
                    If new_ev = "EV�L" Then
                        condition = condition & "��"
                    Else
                        condition = condition & "�~"
                    End If
                End If
                        
                ' �Z��������
                address = ""
                        
                ' �Z���ϐ���`
                ' ���Z��
                now_address = myArray(3, j)
                ' �V�Z��
                new_address = myArray(7, j)
                ' ���V�Z��
                address = now_address & " �` " & new_address
                        
                If i = 1 Then
                    ' �Z���Ɋi�[
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
                    ' �Z���Ɋi�[
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
                        ' �Z���Ɋi�[
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
            Next
        End If
    Next
    
    ' AM�EPM�Efree�̃f�[�^�������ꍇ���b�Z�[�W�\��
    If Eofnum = 3 Then
        MsgBox "���q�l�f�[�^������܂���"
    End If
    
    ' �������
    adoRs.Close
    Set adoRs = Nothing
    db.disconnect
    Set db = Nothing
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

' �\��X�V
Sub Re_UPDATA_Click()
    Dim db As DBManager     ' �f�[�^�x�[�X�N���X
    Dim dbFlg As Boolean    ' �ڑ��t���O
    Dim SQL As String       ' SQL
    Dim result As Long      ' YesNo�{�^���\��
    
    On Error GoTo ErrorTrap
   
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Set db = New DBManager
    dbFlg = db.connect
          
    result = MsgBox("�㏑���ۑ����Ă���낵���ł����H", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        ' AM��PM
        For i = 0 To 1
            ' 13�s�i�z�ԕ\�̌ߑO+free�E�ߌ�+free�j
            For j = 0 To 13
                ' AM
                If i = 0 Then
                    If Range("E" & j * 4 + 4).Value <> "" Then
                        ' SQL��
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
                        
                        ' SQL�̎��s
                        db.execute SQL
                    End If
                ' PM
                Else
                    If Range("Q" & j * 4 + 4).Value <> "" Then
                        ' SQL��
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
                        
                        ' SQL�̎��s
                        db.execute SQL
                    End If
                End If
            Next j
        Next i
    End If
 
    ' �������
    db.disconnect
    Set db = Nothing
Exit Sub
 
ErrorTrap:
    ' �������
    Set db = Nothing
    
    ' �G���[����
    Select Case Err.Number
        ' DB�ڑ��G���[
        Case -2147467259
            MsgBox "�f�[�^�x�[�X�ɐڑ��ł��܂���"
    End Select
End Sub
