Attribute VB_Name = "���q�l���"
Option Explicit

'���q�l���N���A
Sub �N���A_Click()
    ' �Z���̓��e���N���A����B
    Range("I5") = ""    ' ���q�lID
    Range("X9") = ""    ' ���q�l����
    Range("B9") = ""    ' ��]��1
    Range("J9") = ""    ' ��]��2
    Range("Q9") = ""    ' am,pm,free
    Range("S9") = ""    ' �J�n���ԑO
    Range("V9") = ""    ' �J�n���Ԍ�
    Range("I6") = ""    ' ��]�����R
    Range("AE6") = ""   ' ����d�b�ԍ�1
    Range("AI6") = ""   ' ����d�b�ԍ�2
    Range("AN6") = ""   ' ����d�b�ԍ�3
    Range("AE7") = ""   ' �A����d�b�ԍ�1
    Range("AI7") = ""   ' �A����d�b�ԍ�2
    Range("AN7") = ""   ' �A����d�b�ԍ�3
    Range("K12") = ""   ' ���Z��
    Range("K11") = ""   ' ����1
    Range("O11") = ""   ' ����2
    Range("C13") = ""   ' ���K��
    Range("I13") = ""   ' ��ev
    Range("G14") = ""   ' ������
    Range("AM11") = ""  ' ���������
    Range("K17") = ""   ' �V�Z��
    Range("K16") = ""   ' �V��1
    Range("O16") = ""   ' �V��2
    Range("C18") = ""   ' �V�K�w
    Range("I18") = ""   ' �Vev
    Range("G19") = ""   ' �V����
    Range("AM16") = ""  ' �V�������
    Range("AR8") = ""   ' ��t��1
    Range("AV8") = ""   ' ��t��2
    Range("AZ8") = ""   ' ��t��3
    Range("BD8") = ""   ' ��t��4
    Range("AU11") = ""  ' ��t�S����
    Range("AR15") = ""  ' ������1
    Range("AV15") = ""  ' ������2
    Range("AZ15") = ""  ' ������3
    Range("BD15") = ""  ' ������4
    Range("AU18") = ""  ' �����S����
    Range("M21:M69") = ""   ' �ו���
    Range("Z21:Z69") = ""
    Range("AM21:AM69") = ""
    Range("BC21:BC45") = ""
    Range("AY49") = ""
    Range("AY54") = ""
    Range("BC55:BC59") = ""
    Range("AZ73").Value = "=SUM(K71+X71+AK71+AZ71)" ' �|�C���g���v
End Sub

' �Q��
Sub SELECT_Click()
    ' �t�H�[����\������B
    �Q��Form.Show
End Sub

' �ǉ�
Sub INSERT_Click()
    ' ���[�N�V�[�g��\�����ړ�����B
    Worksheets("�V�K�쐬").Visible = True
    Worksheets("�V�K�쐬").Activate
    Range("A1").Select
End Sub

' �X�V
Sub UPDATA_Click()
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Dim db As DBManager         ' �f�[�^�x�[�X�N���X
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    
    
    ' �m�F�{�^����\�����A�m�F�{�^����Yes�̏ꍇ�A�������������l�𒴂��Ă��Ȃ����`�F�b�N����B
    Dim result As Long  ' YesNo�{�^���\��
    
    result = MsgBox("�㏑���ۑ����Ă���낵���ł����H", vbYesNo + vbExclamation + vbDefaultButton2)

    If result = vbYes Then
        If Range("I5").Value = "" Then
            MsgBox "ID���I������Ă��܂���"
            End
        End If
        
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
        
            
            ' Date�֐��ɂ���č����̓��t���擾���A�Z���̓��t������̓��t�ȑO�ł���ꍇ�N�����Z����B
            Dim thisyear As Integer     ' �N
            Dim thismonth As Integer    ' ��
            Dim thisday As Integer      ' ��
            
            thisyear = year(Date)
            thismonth = Month(Date)
            thisday = Day(Date)
            
            If thismonth >= Range("B9").Value Then
                If thisday > Range("J9").Value Then
                    thisyear = year(Date) + 1
                End If
            End If
    
    
            ' ���q�l���V�[�g�̃Z���̒l���擾����UPDATE����SQL����ݒ肵���s����B
            Dim SQL As String   ' SQL
            
            SQL = "UPDATE customers SET name = '" & Range("X9") & "'," & _
            "move_day = '" & thisyear & "-" & Range("B9").Value & "-" & Range("J9").Value & "'," & _
            "meridian = '" & Range("Q9").Value & "',front_time = '" & Range("S9").Value & "'," & _
            "back_time = '" & Range("V9").Value & "',reason = '" & Range("I6").Value & "'," & _
            "home_phone = '" & Range("AE6").Value & "," & Range("AI6").Value & "," & Range("AN6").Value & "'," & _
            "contact_phone = '" & Range("AE7").Value & "," & Range("AI7").Value & "," & Range("AN7").Value & "'," & _
            "now_address = '" & Range("K12").Value & "'," & _
            "now_postalcode = '" & Range("K11").Value & "," & Range("O11").Value & "'," & _
            "now_floors = '" & Range("C13").Value & "',now_ev = '" & Range("I13").Value & "'," & _
            "now_width = '" & Range("G14").Value & "',now_type = '" & Range("AM11").Value & "'," & _
            "new_address = '" & Range("K17").Value & "'," & _
            "new_postalcode = '" & Range("K16").Value & "," & Range("O16").Value & "'," & _
            "new_floors = '" & Range("C18").Value & "',new_ev = '" & Range("I18").Value & "'," & _
            "new_width = '" & Range("G19").Value & "',new_type = '" & Range("AM16").Value & "'," & _
            "reception_day = '1900-''" & Range("AR8").Value & "-" & Range("AV8").Value & " " & Range("AZ8").Value & "" & _
            ":" & Range("BD8").Value & "'':00',reception_name = '" & Range("AU11").Value & "'," & _
            "preview_day = '1900-''" & Range("AR15").Value & "-" & Range("AV15").Value & " " & Range("AZ15").Value & "" & _
            ":" & Range("BD15").Value & "'':00',preview_name = '" & Range("AU18").Value & "'," & _
            "point = '" & Range("AZ73").Value & "' WHERE id = " & Range("I5")
            
            db.execute SQL
        End If
    End If
 

    ' �f�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
    db.disconnect
    Set db = Nothing
Exit Sub
 
ErrorTrap:
    ' �G���[�����������ꍇ�A�f�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
    Set db = Nothing
End Sub

' �폜
Sub DELETE_Click()
    ' �t�H�[����\������B
    �폜Form.Show
End Sub
