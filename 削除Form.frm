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

Private Sub UserForm_Initialize()
    Dim db As DBManager         ' �f�[�^�x�[�X�N���X
    Dim dbFlg As Boolean        ' �ڑ��t���O
    Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
    Dim SQL As String           ' SQL
    Dim FldCount As Integer     ' �t�B�[���h��
    Dim RecCount As Long        ' ���R�[�h��
    Dim myArray() As Variant    ' �Q�ƃ��R�[�h�z��
     
    On Error GoTo ErrorTrap
   
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Set db = New DBManager
    dbFlg = db.connect
    
    ' SQL��
    SQL = "SELECT id,name,DATE_FORMAT(move_day,'%c��%e��'),preview_name," & _
    "DATE_FORMAT(preview_day,'%c��%e�� %H��%i��') FROM customers"
     
    ' SQL�̎��s
    Set adoRs = db.execute(SQL)
    
    ' �t�B�[���h���ƃ��R�[�h�����擾
    FldCount = adoRs.Fields.Count
    RecCount = adoRs.RecordCount
    
    ' ���R�[�h�������ꍇ
    If adoRs.EOF Then
          MsgBox "���q�l�f�[�^������܂���"
          ' �v���O�����̏I��
          End
    Else
        ' �񎟌��z����Ē�`
        ReDim myArray(FldCount - 1, RecCount - 1)
        ' ���R�[�h�Z�b�g�̓��e��ϐ��Ɋi�[
        myArray = adoRs.GetRows
    
        ' ���X�g�{�b�N�X��o�^�i�`�F�b�N�{�b�N�X�E�����I���j
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

' �폜
Private Sub DELETEButton_Click()
    Dim db As DBManager         ' �f�[�^�x�[�X�N���X
    Dim dbFlg As Boolean        ' �ڑ��t���O
    Dim SQL As String           ' SQL
    Dim adoRs As Object         ' ADO���R�[�h�Z�b�g
    Dim FldCount As Integer     ' �t�B�[���h��
    Dim RecCount As Long        ' ���R�[�h��
    Dim myArray() As Variant    ' �Q�ƃ��R�[�h�z��
    Dim result As String        ' YesNo�{�^���\��

    On Error GoTo ErrorTrap
    
    ' DBManager�N���X��db���C���X�^���X�����A�ڑ��������s��
    Set db = New DBManager
    dbFlg = db.connect
    
    result = MsgBox("�f�[�^���폜���Ă���낵���ł����H", vbYesNo + vbExclamation + vbDefaultButton2)
    
    If result = vbYes Then
        ' ���X�g�{�b�N�X�őI���������q�l�f�[�^���폜
        With ���q�l���X�g
            For i = 0 To .ListCount - 1
                If .Selected(i) = True Then
                    ' SQL��
                    SQL = "DELETE FROM customers WHERE id=" & .List(i, 0)
                      
                    ' SQL�̎��s
                    db.execute SQL
                End If
            Next i
        End With
        
        ' �폜��ĕ\������
        ' SQL��
        SQL = "SELECT id,name,DATE_FORMAT(move_day,'%c��%e��'),preview_name," & _
        "DATE_FORMAT(preview_day,'%c��%e�� %H��%i��') FROM customers"
         
        ' SQL�̎��s
        Set adoRs = db.execute(SQL)
        
        ' �t�B�[���h���ƃ��R�[�h�����擾
        FldCount = adoRs.Fields.Count
        RecCount = adoRs.RecordCount
        
        ' ���R�[�h�������ꍇ
        If adoRs.EOF Then
            With ���q�l���X�g
                .Clear
            End With
        Else
            ' �񎟌��z����Ē�`
            ReDim myArray(FldCount - 1, RecCount - 1)
            ' ���R�[�h�Z�b�g�̓��e��ϐ��Ɋi�[
            myArray = adoRs.GetRows
        
            ' ���X�g�{�b�N�X���X�V�i�`�F�b�N�{�b�N�X�E�����I���j
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

' ����{�^��
Private Sub ����_Click()
    Unload Me
End Sub

