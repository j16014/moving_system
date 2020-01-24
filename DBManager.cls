VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private adoCon As ADODB.Connection  ' ADO�R�l�N�V����
Private adoRs As ADODB.Recordset    ' ADO���R�[�h�Z�b�g

' �f�[�^�x�[�X�֐ڑ�
Public Function connect() As Boolean
    Set adoCon = New ADODB.Connection   ' ADO�R�l�N�V�������쐬

    On Error GoTo ErrorTrap
    
    '192.168.11.11 ����
    '10.206.3.155 �w�Z
    '192.168.100.102 �ؑ�
    
    ' ODBC�ڑ�
    adoCon.Open _
        "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
        " SERVER=10.206.3.155;" & _
        " DATABASE=arisan_db;" & _
        " UID=test;" & _
        " PWD=Test123_pass;"
    
    ' �N���C�A���g�T�C�h�J�[�\���ɕύX
    adoCon.CursorLocation = 3
    
    ' �ڑ�����
    connect = True
Exit Function

ErrorTrap:
    ' �ڑ����s
    connect = False

    ' �������
    Set adoCon = Nothing
    
    ' �G���[����
    Select Case Err.Number
        ' DB�ڑ��G���[
        Case -2147467259
            MsgBox "�f�[�^�x�[�X�ɐڑ��ł��܂���"
    End Select
End Function

' �f�[�^�x�[�X�ւ̐ڑ���ؒf
Public Sub disconnect()
    adoCon.Close
    Set adoCon = Nothing
End Sub

' SQL�����s
Public Function execute(SQL As String) As ADODB.Recordset
    Set adoRs = New ADODB.Recordset

    adoRs.Open SQL, adoCon
    
    ' �߂�l
    Set execute = adoRs
End Function