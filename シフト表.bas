Attribute VB_Name = "�V�t�g�\"
Option Explicit

Sub ��]�x����()
    ' �t�H�[����\������B
    �V�t�g�\Form.Show
End Sub

Sub �����̃V�[�g�쐬()
    ' �V�[�g����N�������擾���ϐ��Ɋi�[����B
    Dim d As Date       ' ��
    Dim y As Integer    ' �N
    Dim m As Integer    ' ��
    
    d = Range("I2").Value
    y = year(d)
    m = Month(d)
    

    ' ���𗂌��ɕύX���A�����̃V�[�g����ݒ肵����A�����̃V�[�g�����ɍ쐬����Ă��邩�𔻒肷��B
    Dim NextSheet As String     ' �����̃V�[�g��
    Dim SameSheet As Boolean    ' �����V�[�g����t���O
        
    If m = 12 Then
        y = y + 1
        m = 1
    Else
        m = m + 1
    End If
    
    NextSheet = y & "." & m
    SameSheet = exist_check(NextSheet, "ws")
    
    
    ' �����̃V�[�g�����݂���ꍇ�́A���b�Z�[�W��\������B
    ' �����̃V�[�g�����݂��Ȃ��ꍇ�͐V�K�V�[�g���쐬���A�V�[�g���̕ύX�E�Z���̓��e�̕ύX�E�Z���̃N���A���s���B
    If SameSheet = True Then
        MsgBox "�����̃V�[�g�͊��ɍ쐬����Ă��܂�"
    Else
        Worksheets(ActiveSheet.name).Copy After:=ActiveSheet
        ActiveSheet.name = NextSheet
             
        Range("I2").Value = DateAdd("m", 1, d)
        
        Range(Cells(6, 9), Cells(56, 39)).ClearContents
    End If
End Sub

Sub �V�t�g�\�R�s�[()
    ' �R�s�[���u�b�N���ƌ��݃}�N�������s���Ă���Excel�u�b�N����R�s�[��u�b�N����ݒ肷��B
    ' �R�s�[���u�b�N���́u���_�V�t�g������.xlsm�v�A�R�s�[��u�b�N���́u���_�V�X�e��.xlsm�v�ł���B
    Dim SourceBook As String ' �R�s�[���u�b�N��
    Dim TargetBook As String ' �R�s�[��u�b�N��
    
    SourceBook = "���_�V�t�g������.xlsm"
    TargetBook = Application.ThisWorkbook.name
    
    
    ' �R�s�[���u�b�N���J����Ă��邩�A�R�s�[���u�b�N���ǂݎ���p�ł��邩���m�F���A�����𖞂������ꍇ���b�Z�[�W��\������B
    Dim SameBook As Boolean ' �����u�b�N����t���O
    
    SameBook = exist_check(SourceBook, "wb")
    
    If SameBook = False Then
        MsgBox SourceBook & " ���J����Ă��܂���" & vbCrLf & "�u�b�N���J���Ă�����s���Ă�������"
        Exit Sub
    End If
    
    If Workbooks(SourceBook).ReadOnly = True Then
        MsgBox SourceBook & " ���ǂݎ���p�ɂȂ��Ă��܂�" & vbCrLf & "�ҏW�\�ɂ��Ď��s���Ă�������"
        Exit Sub
    End If
    
    
    ' �R�s�[���u�b�N�ƃR�s�[��u�b�N�ŃZ���͈̔͂��w�肵�A�R�s�[�������s���B
    ' �R�s�[������A�R�s�[���[�h����������B
    Workbooks(SourceBook).Worksheets("�Ζ��X�P�W���[���C��").Range("E6:AI15").Copy
    Workbooks(TargetBook).Worksheets(ActiveSheet.name).Range("I6:AM15").PasteSpecial (xlPasteValues)
    
    Application.CutCopyMode = False
End Sub
