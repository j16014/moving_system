VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �V�t�g�\Form 
   Caption         =   "��]�x����"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4005
   OleObjectBlob   =   "�V�t�g�\Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�V�t�g�\Form"
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
    ' �}�X�^��������擾���A�t�H�[���̃R���{�{�b�N�X�ɐݒ肷��B
    Dim Lastrow As Integer  ' �}�X�^�̍ŏI�s
    
    NameBox.Clear
        Lastrow = Worksheets("����}�X�^").Cells(Rows.Count, 3).End(xlUp).Row
    For i = 4 To Lastrow
        NameBox.AddItem Worksheets("����}�X�^").Range("C" & i).Value
    Next i


    ' �����̖������擾����B
    Dim d As Date           ' ��
    Dim sLastday As Date    ' �����̖���
    
    d = Range("I2").Value
        
    sLastday = Format(DateSerial(year(d), Month(d) + 1, 0), "d")
    
    
    ' �����̓��ɂ��ɂ���ă`�F�b�N�{�b�N�X���\���ɂ���B
    If sLastday = 28 Then
        Label29.Enabled = False
        Label30.Enabled = False
        Label31.Enabled = False
        CheckBox29.Enabled = False
        CheckBox30.Enabled = False
        CheckBox31.Enabled = False
    ElseIf sLastday = 29 Then
        Label30.Enabled = False
        Label31.Enabled = False
        CheckBox30.Enabled = False
        CheckBox31.Enabled = False
    ElseIf sLastday = 30 Then
        Label31.Enabled = False
        CheckBox31.Enabled = False
    End If
End Sub

' �������I�����ꂽ�珈��
Private Sub NameBox_Change()
    ' �����R���{�{�b�N�X�̓��e���ύX���ꂽ���A�ύX���ꂽ���e�����ɊY������Z������l���擾���A31�̃`�F�b�N�{�b�N�X�̓��e��ύX����B
    Dim Index As Integer    ' �R���{�{�b�N�X�̔ԍ�
    
    Index = NameBox.ListIndex

    If Index <> -1 Then
            Index = Index + 10
        
        For i = 1 To 31
            If Cells(6 + Index, 8 + i).Value = "��" Then
                Me.Controls("CheckBox" & i) = True
            Else
                Me.Controls("CheckBox" & i) = False
            End If
        Next i
    End If
End Sub

' �S�đI���{�^���������đS�ẴR���{�{�b�N�X��I����Ԃɂ���
Private Sub SelectionButton_Click()
    ' �S�đI���{�^�������������A�S�ẴR���{�{�b�N�X��I����Ԃɂ��鏈�����s���B
    For i = 1 To 31
        Me.Controls("CheckBox" & i) = True
    Next
End Sub

' �S�ĉ����{�^���������đS�ẴR���{�{�b�N�X�̑I������������
Private Sub ReleaseButton_Click()
    ' �S�ĉ����{�^�������������A�S�ẴR���{�{�b�N�X�̑I�����������鏈�����s���B
    For i = 1 To 31
        Me.Controls("CheckBox" & i) = False
    Next
End Sub

' �����{�^���N���b�N
Private Sub CompleteButton_Click()
    ' �����R���{�{�b�N�X�̓��e���擾����B�R���{�{�b�N�X���I������Ă��Ȃ��ꍇ�̓��b�Z�[�W��\������B
    ' �R���{�{�b�N�X���I������Ă���ꍇ�̓V�[�g����N�������擾���ϐ��Ɋi�[���A�挎�̃V�[�g����ݒ肷��B
    Dim Index As Integer    ' �R���{�{�b�N�X�̃C���f�b�N�X
    Dim d As Date               ' ��
    Dim LastMonth As Date       ' �挎
    Dim y As Integer            ' �N
    Dim m As Integer            ' ��
    Dim LastSheet As String     ' �挎�̃V�[�g��
    
    Index = NameBox.ListIndex
    
    If Index = -1 Then
        MsgBox "�������I������Ă��܂���"
    Else
        Index = Index + 10

        d = Range("I2").Value

        LastMonth = DateAdd("m", -1, d)
    
        y = year(LastMonth)
        m = Month(LastMonth)
        
        LastSheet = y & "." & m
        
        
        ' Excel�̊֐��ƂȂ镶�����ݒ肷��B
        ' �֐��͐挎�̃V�[�g�����݂���ꍇ�ɂ��̖������擾����֐��ł���B
        Dim s As String ' �挎�̖���������
        
        s = "IF(DAY(EOMONTH(" & CStr(LastSheet) & "!I3,0))=28," & CStr(LastSheet) & "!AJ" & 6 + Index & ",IF(DAY(EOMONTH(" & CStr(LastSheet) & "!I3,0))=29," & CStr(LastSheet) & "!AK" & 6 + Index & ",IF(DAY(EOMONTH(" & CStr(LastSheet) & "!I3,0))=30," & CStr(LastSheet) & "!AL" & 6 + Index & "," & CStr(LastSheet) & "!AM" & 6 + Index & ")))"
    
    
        ' 31�����������J��Ԃ��A�`�F�b�N�{�b�N�X���I������Ă��邩�ɂ���ď����𕪊򂳂���B
        ' �`�F�b�N�{�b�N�X���I������Ă���ꍇ�A28���ȑO�ƈȍ~�ňقȂ�֐����Z���ɐݒ肷��B
        ' �`�F�b�N�{�b�N�X���I������Ă��Ȃ��ꍇ�A1���A28���ȍ~�ŏ����𕪊򂳂���B1���̏ꍇ�A�挎�̃V�[�g�����݂��邩���m�F�����݂��邩�ǂ����ɂ���ĈقȂ�֐����Z���ɐݒ肷��B
        '28���ȍ~�̏ꍇ�����l�Ɋ֐����Z���ɐݒ肷��B
        Dim SameSheet As Boolean    ' �����V�[�g����t���O
        
        For i = 1 To 31
            If Me.Controls("CheckBox" & i) = True Then
                If i > 28 Then
                    Cells(6 + Index, 8 + i).Value = "=IF(" & Cells(3, 8 + i).address & "="""","""",""��"")"
                Else
                    Cells(6 + Index, 8 + i).Value = "��"
                End If
            ElseIf Me.Controls("CheckBox" & i) = False Then
                If i = 1 Then
                    SameSheet = exist_check(LastSheet, "ws")
                    
                    If SameSheet = True Then
                        Cells(6 + Index, 8 + i).Value = _
                        "=IF(OR(" & s & "=""�x""," & s & "=""��""," & s & "=""AM""," & s & "=""PM""),1," & s & "+1)"
                    Else
                        Cells(6 + Index, 8 + i).Value = 1
                    End If
                ElseIf i > 28 Then
                    Cells(6 + Index, 8 + i).Value = _
                    "=IF(" & Cells(3, 8 + i).address & "="""",""""," & _
                    "IF(OR(" & Cells(6 + Index, 7 + i).address & "=""�x""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""��""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""AM""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""PM""),1," & Cells(6 + Index, 7 + i).address & "+1))"
                Else
                    Cells(6 + Index, 8 + i).Value = _
                    "=IF(OR(" & Cells(6 + Index, 7 + i).address & "=""�x""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""��""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""AM""," & _
                    "" & Cells(6 + Index, 7 + i).address & "=""PM""),1," & Cells(6 + Index, 7 + i).address & "+1)"
                End If
            End If
        Next i
    End If
End Sub

' ����{�^��
Private Sub ����_Click()
    ' �t�H�[�������B
    Unload Me
End Sub

