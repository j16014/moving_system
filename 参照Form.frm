VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �Q��Form 
   Caption         =   "�Q��"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7185
   OleObjectBlob   =   "�Q��Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�Q��Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' �ϐ�
Dim i As Integer
Dim j As Integer
Dim k As Integer

' �N�����ݒ�
Private Sub UserForm_Initialize()
    ' �e�L�X�g�{�b�N�X�����������A�ҏW�s�ɐݒ肷��B
    For i = 1 To 5
        Controls("TextBox" & i).Value = ""
        Controls("TextBox" & i).Locked = True
    Next i
    
    
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Dim db As DBManager     ' �f�[�^�x�[�X�N���X
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    
    
    ' SQL�������s���A���R�[�h�Z�b�g����t�B�[���h���ƃ��R�[�h�����擾����B
    Dim SQL As String           ' SQL
    Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
    Dim FldCount As Integer     ' �t�B�[���h��
    Dim RecCount As Long        ' ���R�[�h��
    
    SQL = "SELECT id,name,DATE_FORMAT(move_day,'%c��%e��'),preview_name," & _
    "DATE_FORMAT(preview_day,'%c��%e�� %H��%i��') FROM customers"
    
    Set adoRs = db.execute(SQL)

    FldCount = adoRs.Fields.Count
    RecCount = adoRs.RecordCount
    
    
    ' ���R�[�h�����݂��Ȃ��ꍇ�́A���b�Z�[�W��\������B
    ' ���R�[�h�����݂���ꍇ�́A���R�[�h�Z�b�g��z��Ɋi�[���A���X�g�{�b�N�X�ɓo�^����B
    Dim myArray() As Variant    ' �Q�ƃ��R�[�h�z��
    
    If adoRs.EOF Then
        MsgBox "���q�l�f�[�^������܂���"
        End
    Else
        ReDim myArray(FldCount - 1, RecCount - 1)
        myArray = adoRs.GetRows
    
        With ���q�l���X�g
            .Clear
            .ColumnCount = 5
            .ColumnWidths = "30;70;70;70"
            .Column = myArray
        End With
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

' �Q��
Private Sub SELECTButton_Click()
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Dim db As DBManager     ' �f�[�^�x�[�X�N���X
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    
    
    ' �Z���̓��e���N���A������ASQL�������s���ă��R�[�h�Z�b�g����t�B�[���h���ƃ��R�[�h�����擾����B
    Dim SQL As String           ' SQL
    Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
    Dim FldCount As Integer     ' �t�B�[���h��
    Dim RecCount As Long        ' ���R�[�h��
    
    Call �N���A_Click
    
    SQL = "SELECT name,DATE_FORMAT(move_day,'%c,%e'),meridian,front_time,back_time,reason," & _
    "home_phone,contact_phone,now_address,now_postalcode,now_floors,now_ev,now_width,now_type," & _
    "new_address,new_postalcode,new_floors,new_ev,new_width,new_type," & _
    "DATE_FORMAT(reception_day,'%c,%e,%H,%i'),reception_name," & _
    "DATE_FORMAT(preview_day,'%c,%e,%H,%i'),preview_name,point " & _
    "FROM customers WHERE id = '" & TextBox1.Value & "'"

    Set adoRs = db.execute(SQL)
    
    FldCount = adoRs.Fields.Count
    RecCount = adoRs.RecordCount
    
    
    ' ���R�[�h�����݂��Ȃ��ꍇ�̓��b�Z�[�W��\������B
    ' ���R�[�h�����݂���ꍇ�̓��R�[�h�Z�b�g��z��Ɋi�[���A�z�񂩂�e�Z���ɒl��}������B
    Dim myArray() As Variant    ' �Q�ƃ��R�[�h�z��
    Dim splitArray() As String  ' ��؂�z��
    
    If adoRs.EOF Then
        MsgBox "���q�l�f�[�^���I������Ă��܂���"
    Else
        ReDim myArray(FldCount - 1, RecCount - 1)
        myArray = adoRs.GetRows

        Range("I5") = TextBox1.Value        ' ���q�lID
        Range("X9") = myArray(0, 0)         ' ���q�l����
        splitArray = Split(myArray(1, 0), ",")
        Range("B9") = splitArray(0)         ' ��]��1
        Range("J9") = splitArray(1)         ' ��]��2
        Range("Q9") = myArray(2, 0)         ' am,pm,free
        Range("S9") = myArray(3, 0)         ' �J�n���ԑO
        Range("V9") = myArray(4, 0)         ' �J�n���Ԍ�
        Range("I6") = myArray(5, 0)         ' ��]�����R
        splitArray = Split(myArray(6, 0), ",")
        Range("AE6") = splitArray(0)        ' ����d�b�ԍ�1
        Range("AI6") = splitArray(1)        ' ����d�b�ԍ�2
        Range("AN6") = splitArray(2)        ' ����d�b�ԍ�3
        splitArray = Split(myArray(7, 0), ",")
        Range("AE7") = splitArray(0)        ' �A����d�b�ԍ�1
        Range("AI7") = splitArray(1)        ' �A����d�b�ԍ�2
        Range("AN7") = splitArray(2)        ' �A����d�b�ԍ�3
        Range("K12") = myArray(8, 0)        ' ���Z��
        splitArray = Split(myArray(9, 0), ",")
        Range("K11") = splitArray(0)        ' ����1
        Range("O11") = splitArray(1)        ' ����2
        Range("C13") = myArray(10, 0)       ' ���K��
        Range("I13") = myArray(11, 0)       ' ��ev
        Range("G14") = myArray(12, 0)       ' ������
        Range("AM11") = myArray(13, 0)      ' ���������
        Range("K17") = myArray(14, 0)       ' �V�Z��
        splitArray = Split(myArray(15, 0), ",")
        Range("K16") = splitArray(0)        ' �V��1
        Range("O16") = splitArray(1)        ' �V��2
        Range("C18") = myArray(16, 0)       ' �V�K�w
        Range("I18") = myArray(17, 0)       ' �Vev
        Range("G19") = myArray(18, 0)       ' �V����
        Range("AM16") = myArray(19, 0)      ' �V�������
        splitArray = Split(myArray(20, 0), ",")
        Range("AR8") = splitArray(0)        ' ��t��1
        Range("AV8") = splitArray(1)        ' ��t��2
        Range("AZ8") = splitArray(2)        ' ��t��3
        Range("BD8") = splitArray(3)        ' ��t��4
        Range("AU11") = myArray(21, 0)      ' ��t�S����
        splitArray = Split(myArray(22, 0), ",")
        Range("AR15") = splitArray(0)       ' ������1
        Range("AV15") = splitArray(1)       ' ������2
        Range("AZ15") = splitArray(2)       ' ������3
        Range("BD15") = splitArray(3)       ' ������4
        Range("AU18") = myArray(23, 0)      ' �����S����
        Range("AZ73") = "=SUM(K71+X71+AK71+AZ71)+" & myArray(24, 0) '�|�C���g
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

Private Sub ���q�l���X�g_Click()
    ' ���X�g�{�b�N�X��ŃN���b�N�������ڂ��e�L�X�g�{�b�N�X�Ɋi�[����B
    Dim sIndex
    
    sIndex = ���q�l���X�g.ListIndex
    TextBox1.Text = ���q�l���X�g.List(sIndex, 0)
    TextBox2.Text = ���q�l���X�g.List(sIndex, 1)
    TextBox3.Text = ���q�l���X�g.List(sIndex, 2)
    TextBox4.Text = ���q�l���X�g.List(sIndex, 3)
    TextBox5.Text = ���q�l���X�g.List(sIndex, 4)
End Sub

' ����{�^��
Private Sub ����_Click()
    ' �t�H�[�������B
    Unload Me
End Sub
