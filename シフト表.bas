Attribute VB_Name = "�V�t�g�\"
Option Explicit

Sub ��]�x����()
    �V�t�g�\Form.Show
End Sub

Sub �����̃V�[�g�쐬()
    Dim d As Date               ' �V�[�g�̓���
    Dim y As Integer            ' �N�ϐ�
    Dim m As Integer            ' ���ϐ�
    Dim NextSheet As String     ' �쐬�����V�[�g��
    Dim SameSheet As Boolean    ' �����V�[�g����t���O
    Dim ws As Worksheet         ' ���[�N�V�[�g
    Dim SheetName As String     ' ���[�N�V�[�g��
    Dim NextMonth As Date       ' ����
    
    ' I2�̔N�������擾
    d = Range("I2").Value
    y = year(d)
    m = Month(d)
    
    ' 12���Ȃ�N�x��ς���
    If m = 12 Then
        y = y + 1
        m = 1
    Else
        m = m + 1
    End If
    
    ' �����̃V�[�g���쐬
    NextSheet = y & "." & m
    
    ' �����̃V�[�g�����݂����烁�b�Z�[�W�A���݂��Ȃ���Η����̃V�[�g�쐬
    SameSheet = False
    
    ' �u�b�N���̃��[�N�V�[�g�����擾
    For Each ws In Sheets
        If ws.name = NextSheet Then
            ' ���݂���
            SameSheet = True
        End If
    Next
    
    If SameSheet = True Then
        MsgBox "�����̃V�[�g�͊��ɍ쐬����Ă��܂�"
        Exit Sub
    Else
        ' ���[�N�V�[�g�����擾
        SheetName = ActiveSheet.name
        
        ' �����̃V�[�g�쐬
        Worksheets(SheetName).Copy After:=ActiveSheet
        ActiveSheet.name = NextSheet
             
        ' ���������
        NextMonth = DateAdd("m", 1, d)
        Range("I2").Value = NextMonth
        
        ' �Z���̓��e���N���A����
        Range(Cells(6, 9), Cells(56, 39)).ClearContents
    End If
End Sub

Sub �V�t�g�\�R�s�[()
    
    Dim SourceFile As String    ' �R�s�[���t�@�C����
    Dim TargetFile As String    ' �R�s�[��t�@�C����
    Dim SameFile As Boolean     ' �����t�@�C������t���O
    Dim wb As Workbook          ' ���[�N�u�b�N
    
    ' �R�s�[���t�@�C����
    SourceFile = "���_�V�t�g������.xlsm"
    ' �R�s�[��t�@�C����
    TargetFile = Application.ThisWorkbook.name
    
    ' �t�@�C�������ɊJ���Ă��邩�`�F�b�N
    SameFile = False
    
    For Each wb In Workbooks
        If wb.name = SourceFile Then
            ' �J���Ă���
            SameFile = True
        End If
    Next wb
    
    ' �R�s�[���t�@�C�����J���Ă��邩�`�F�b�N
    If SameFile = False Then
        MsgBox SourceFile & " ���J����Ă��܂���" & vbCrLf & "�t�@�C�����J���Ă�����s���Ă�������"
        Exit Sub
    End If
    
    ' �ǂݎ���p�`�F�b�N
    If Workbooks(SourceFile).ReadOnly = True Then
        MsgBox SourceFile & " ���ǂݎ���p�ɂȂ��Ă��܂�" & vbCrLf & "�ҏW�\�ɂ��Ď��s���Ă�������"
        Exit Sub
    End If
    
    ' �R�s�[����
    Workbooks(SourceFile).Worksheets("�Ζ��X�P�W���[���C��").Range("E6:AI15").Copy
    Workbooks(TargetFile).Worksheets(ActiveSheet.name).Range("I6:AM15").PasteSpecial (xlPasteValues)
    
    ' �R�s�[���[�h����
    Application.CutCopyMode = False
End Sub
