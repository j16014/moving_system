Attribute VB_Name = "�V�K�쐬"
Option Explicit

' �ǉ�
Sub NewINSERT_Click()
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Dim db As DBManager     ' �f�[�^�x�[�X�N���X
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    
    
    ' �m�F�{�^����\�����AYes�̏ꍇ�A�������������l�𒴂��Ă��Ȃ����`�F�b�N����B
    Dim result As Long  ' YesNo�{�^���\��
    
    result = MsgBox("�f�[�^��ǉ����Ă���낵���ł����H", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        If Len(Range("X9").Value) <= 20 And _
        Len(Range("B9").Value) <= 2 And Len(Range("J9").Value) <= 2 And _
        Len(Range("Q9").Value) <= 4 And Len(Range("S9").Value) <= 10 And Len(Range("V9").Value) <= 10 And _
        Len(Range("I6").Value) <= 255 And _
        Len(Range("AE6").Value) + Len(Range("AI6").Value) + Len(Range("AN6").Value) <= 13 And _
        Len(Range("AE7").Value) + Len(Range("AI7").Value) + Len(Range("AN7").Value) <= 13 And _
        Len(Range("K12").Value) <= 100 And Len(Range("K11").Value) + Len(Range("O11").Value) <= 7 And _
        Len(Range("C13").Value) <= 3 And Len(Range("I13").Value) <= 3 And _
        Len(Range("G14").Value) <= 1 And Len(Range("AM11").Value) <= 10 And _
        Len(Range("K17").Value) <= 100 And Len(Range("K16").Value) + Len(Range("O16").Value) <= 7 And _
        Len(Range("C18").Value) <= 3 And Len(Range("I18").Value) <= 3 And _
        Len(Range("G19").Value) <= 1 And Len(Range("AM16").Value) <= 10 And _
        Len(Range("AR8").Value) <= 2 And Len(Range("AV8").Value) <= 2 And _
        Len(Range("AZ8").Value) <= 2 And Len(Range("BD8").Value) <= 2 And Len(Range("AU11").Value) <= 20 And _
        Len(Range("AR15").Value) <= 2 And Len(Range("AV15").Value) <= 2 And _
        Len(Range("AZ15").Value) <= 2 And Len(Range("BD15").Value) <= 2 And Len(Range("AU18").Value) <= 20 And _
        Len(Range("AZ73").Value) <= 5 _
        Then
            
            
            ' Date�֐��ɂ���č����̓��t���擾���A�Z���̓��t�������̓��t�ȍ~�������ꍇ�A�N�����Z����B
            Dim thisyear As Integer     ' �N
            Dim thismonth As Integer    ' ��
            Dim thisday As Integer      ' ��
            
            thisyear = year(Date)
            thismonth = Month(Date)
            thisday = Day(Date)
            
            If thismonth >= Range("B9").Value Then
                If thisday >= Range("J9").Value Then
                    thisyear = year(Date) + 1
                End If
            End If
            
            
            ' �Z���ɓ��t�����͂���Ă��Ȃ��ꍇ�̓f�t�H���g�l��ݒ肷��B���͂���Ă���ꍇ�́A�Z���̒l���擾���ĕϐ��Ɋi�[����B
            Dim move_day As String      ' ���z��
            Dim reception_day As String ' ��t��
            Dim preview_day As String   ' ������
            
            move_day = "1900-01-01"
            reception_day = "1900-01-01 01:01:00"
            preview_day = "1900-01-01 01:01:00"
            
            If Range("B9").Value <> "" And Range("J9").Value <> "" Then
                move_day = thisyear & "-" & Range("B9").Value & "-" & Range("J9").Value
            End If
            
            If Range("AR8").Value <> "" And Range("AV8").Value <> "" And Range("AZ8").Value <> "" And Range("BD8").Value <> "" Then
                reception_day = "1900-" & Range("AR8").Value & "-" & Range("AV8").Value & " " & _
                "" & Range("AZ8").Value & ":" & Range("BD8").Value & ":00"
            End If
            
            If Range("AR15").Value <> "" And Range("AV15").Value <> "" And Range("AZ15").Value <> "" And Range("BD15").Value <> "" Then
                preview_day = "1900-" & Range("AR15").Value & "-" & Range("AV15").Value & " " & _
                "" & Range("AZ15").Value & ":" & Range("BD15").Value & ":00"
            End If
            
            
            ' ���q�l�V�[�g�̃Z���̒l���擾���AINSERT����SQL����ݒ肵�Ď��s����B
            Dim SQL As String   ' SQL
            
            SQL = "INSERT INTO customers (name,move_day,meridian,front_time,back_time,reason,home_phone,contact_phone," & _
            "now_address,now_postalcode,now_floors,now_ev,now_width,now_type," & _
            "new_address,new_postalcode,new_floors,new_ev,new_width,new_type," & _
            "reception_day,reception_name,preview_day,preview_name,point,start_time1,start_time2,start_time3," & _
            "plan,difficulty,truck,driver,assistant1,assistant2,assistant3,assistant4) " & _
            " VALUES('" & Range("X9").Value & "','" & move_day & "'," & _
            "'" & Range("Q9").Value & "','" & Range("S9").Value & "','" & Range("V9").Value & "'," & _
            "'" & Range("I6").Value & "'," & _
            "'" & Range("AE6").Value & "," & Range("AI6").Value & "," & Range("AN6").Value & "'," & _
            "'" & Range("AE7").Value & "," & Range("AI7").Value & "," & Range("AN7").Value & "'," & _
            "'" & Range("K12").Value & "','" & Range("K11").Value & "," & Range("O11").Value & "'," & _
            "'" & Range("C13").Value & "','" & Range("I13").Value & "','" & Range("G14").Value & "'," & _
            "'" & Range("AM11").Value & "'," & _
            "'" & Range("K17").Value & "','" & Range("K16").Value & "," & Range("O16").Value & "'," & _
            "'" & Range("C18").Value & "','" & Range("I18").Value & "','" & Range("G19").Value & "'," & _
            "'" & Range("AM16").Value & "'," & _
            "'" & reception_day & "','" & Range("AU11").Value & "'," & _
            "'" & preview_day & "','" & Range("AU18").Value & "'," & _
            "'" & Range("AZ73").Value & "','-','-','-','-','-','-','-','-','-','-','-' )"
            
            db.execute SQL
         

            ' �f�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
            db.disconnect
            Set db = Nothing
            
            
            ' �Z���̓��e���N���A���A���q�l���V�[�g�ɑJ�ڂ�����A�V�K�쐬�V�[�g���\���ɂ���B
            Call �N���A_Click
            
            Worksheets("���q�l���").Activate
            Worksheets("�V�K�쐬").Visible = False
        Else
            ' �������������l���z���Ă���ꍇ�A�f�[�^�x�[�X�I�u�W�F�N�g�̉���������s�����b�Z�[�W��\������B
            Set db = Nothing
            MsgBox "���������I�[�o�[���Ă��܂�"
        End If
    End If
Exit Sub
     
ErrorTrap:
    ' �G���[�����������ꍇ�A�f�[�^�x�[�X�I�u�W�F�N�g�̉���������s��
    Set db = Nothing
End Sub

' ����
Sub ExitINSERT_Click()
    ' �m�F�{�^����\�����AYes�̏ꍇ�̓Z���̓��e���N���A���Ă��q�l���V�[�g�ɑJ�ڂ���B�����ĐV�K�쐬�V�[�g���\���ɂ���B
    Dim result As Long  ' YesNo�{�^���\��

    result = MsgBox("�V�K�쐬����Ă���낵���ł����H", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        Call �N���A_Click
        
        Range("A1").Select
        Worksheets("���q�l���").Activate
        Worksheets("�V�K�쐬").Visible = False
    End If
End Sub

