VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �폜Form 
   Caption         =   "�폜"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7260
   OleObjectBlob   =   "�폜Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�폜Form"
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
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Dim db As DBManager     ' �f�[�^�x�[�X�N���X
    
    On Error GoTo ErrorTrap
   
    Set db = New DBManager
    db.connect
    
    
    ' �m�F�{�^����\�����A�m�F�{�^����Yes�̏ꍇ�ASQL�������s���A���R�[�h�Z�b�g����t�B�[���h���ƃ��R�[�h�����擾����B
    Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
    Dim SQL As String           ' SQL
    Dim FldCount As Integer     ' �t�B�[���h��
    Dim RecCount As Long        ' ���R�[�h��

    SQL = "SELECT id,name,DATE_FORMAT(move_day,'%c��%e��'),preview_name," & _
    "DATE_FORMAT(preview_day,'%c��%e�� %H��%i��') FROM customers"
     
    Set adoRs = db.execute(SQL)
    
    FldCount = adoRs.Fields.Count
    RecCount = adoRs.RecordCount
    
    
    ' ���R�[�h�������ꍇ�̓��b�Z�[�W��\������B
    ' ���R�[�h�����݂���ꍇ�̓��R�[�h�Z�b�g��z��Ɋi�[���A���X�g�{�b�N�X�ɓo�^����B���X�g�{�b�N�X�͕����I���\�ɂ���B
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
            .ListStyle = fmListStyleOption
            .MultiSelect = fmMultiSelectMulti
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

' �폜
Private Sub DELETEButton_Click()
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s���B
    Dim db As DBManager     ' �f�[�^�x�[�X�N���X
    
    On Error GoTo ErrorTrap
    
    Set db = New DBManager
    db.connect
    

    ' �m�F�{�^����\�����AYes�̏ꍇ�̓��X�g�{�b�N�X�őI������Ă��鍀�ڂ��Ƃ�DELETE�����s����B
    Dim result As String        ' YesNo�{�^���\��
    Dim SQL As String   ' SQL
    
    result = MsgBox("�f�[�^���폜���Ă���낵���ł����H", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        With ���q�l���X�g
            For i = 0 To .ListCount - 1
                If .Selected(i) = True Then
                    SQL = "DELETE FROM customers WHERE id=" & .List(i, 0)
                    
                    db.execute SQL
                End If
            Next i
        End With
        

        ' ���R�[�h�폜��t�H�[���ɍĕ\�����s���B
        ' SQL�������s���A���R�[�h�Z�b�g����t�B�[���h���ƃ��R�[�h�����擾����B
        Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
        Dim FldCount As Integer     ' �t�B�[���h��
        Dim RecCount As Long        ' ���R�[�h��
        
        SQL = "SELECT id,name,DATE_FORMAT(move_day,'%c��%e��'),preview_name," & _
        "DATE_FORMAT(preview_day,'%c��%e�� %H��%i��') FROM customers"
         
        Set adoRs = db.execute(SQL)
        
        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
        
        
        ' ���R�[�h�����݂��Ȃ��ꍇ�́A���X�g�{�b�N�X���N���A����B
        ' ���R�[�h�����݂���ꍇ�̓��R�[�h�Z�b�g��z��Ɋi�[���A���X�g�{�b�N�X�ɓo�^����B���X�g�{�b�N�X�͕����I���\�ɂ���B
        Dim myArray() As Variant    ' �Q�ƃ��R�[�h�z��
        
        If adoRs.EOF Then
            With ���q�l���X�g
                .Clear
            End With
        Else
            ReDim myArray(FldCount - 1, RecCount - 1)
            myArray = adoRs.GetRows
        
            With ���q�l���X�g
                .Clear
                .ColumnCount = 5
                .ColumnWidths = "30;70;70;70"
                .Column = myArray
                .ListStyle = fmListStyleOption
                .MultiSelect = fmMultiSelectMulti
            End With
        End If


        ' �������
        adoRs.Close
        Set adoRs = Nothing
    End If
    
    
    ' �f�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
    db.disconnect
    Set db = Nothing
Exit Sub

ErrorTrap:
    ' �G���[�����������ꍇ�A�f�[�^�x�[�X�I�u�W�F�N�g�̉���������s���B
    Set db = Nothing
End Sub

' ����{�^��
Private Sub ����_Click()
    ' �t�H�[�������B
    Unload Me
End Sub

