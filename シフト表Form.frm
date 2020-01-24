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
    Dim d As Date           ' �V�[�g�̓����ϐ�
    Dim sLast As Date       ' ����1���̑O��
    Dim sLastday As Date    ' �����̖���
    Dim Lastrow As Integer  ' �}�X�^�̍ŏI�s
    
    ' �R���{�{�b�N�X�����l �}�X�^����擾
    NameBox.Clear
        Lastrow = Worksheets("����}�X�^").Cells(Rows.Count, 3).End(xlUp).Row
    For i = 4 To Lastrow
        NameBox.AddItem Worksheets("����}�X�^").Range("C" & i).Value
    Next

    d = Range("I2").Value
        
    ' �����P���̑O�����擾
    sLast = DateSerial(year(d), Month(d) + 1, 0)
    
    ' �����̓��݂̂��擾
    sLastday = Format(sLast, "d")
    
    ' �����̓����ɂ���ă��x���ƃ{�^�����\��
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

' �����{�^���N���b�N
Private Sub CompleteButton_Click()
    Dim s As String             ' �挎�̖���������
    Dim str As String           ' �挎������
    Dim Index As Integer        ' �R���{�{�b�N�X�̔ԍ��ϐ�
    Dim d As Date               ' �V�[�g�̓����ϐ�
    Dim y As Integer            ' �N�ϐ�
    Dim m As Integer            ' ���ϐ�
    Dim LastMonth As Date       ' �挎
    Dim SameSheet As Boolean    ' �����V�[�g����t���O
    Dim ws As Worksheet         ' ���[�N�V�[�g
    
    ' �����I������
    Index = NameBox.ListIndex
    
    ' �������I������Ă��Ȃ���΃G���[�A����Ȃ�Index + 10
    If Index = -1 Then
        MsgBox "�������I������Ă��܂���"
        Exit Sub
    Else
        Index = Index + 10

        ' �V�t�g�\�̓����擾
        d = Range("I2").Value
        ' ����������i�挎�ɂ���j
        LastMonth = DateAdd("m", -1, d)
    
        ' �挎�̔N�ƌ�
        y = year(LastMonth)
        m = Month(LastMonth)
        
        ' �挎�̃V�[�g��
        str = y & "." & m
        
        ' �挎�̖����V���[�g�J�b�gs
        s = "IF(DAY(EOMONTH(" & CStr(str) & "!I3,0))=28," & CStr(str) & "!AJ" & 6 + Index & "," & _
        "IF(DAY(EOMONTH(" & CStr(str) & "!I3,0))=29," & CStr(str) & "!AK" & 6 + Index & "," & _
        "IF(DAY(EOMONTH(" & CStr(str) & "!I3,0))=30," & CStr(str) & "!AL" & 6 + Index & "," & _
        "" & CStr(str) & "!AM" & 6 + Index & ")))"
    
        ' �R���{�{�b�N�X�̒l�ɂ�镪��
        For i = 1 To 31
            ' 1���͑O���̍ŏI���𔽉f�@28���ȍ~���������͋�
            If Me.Controls("CheckBox" & i) = True Then
                If i > 28 Then
                    Cells(6 + Index, 8 + i).Value = "=IF(" & Cells(3, 8 + i).address & "="""","""",""��"")"
                Else
                    Cells(6 + Index, 8 + i).Value = "��"
                End If
            ElseIf Me.Controls("CheckBox" & i) = False Then
                If i = 1 Then
                    SameSheet = False
                    
                    ' �u�b�N���ɐ挎�̃V�[�g�����݂��邩�m�F
                    For Each ws In Sheets
                        If ws.name = str Then
                            ' ���݂���
                            SameSheet = True
                        End If
                    Next
                    
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
        Next
    End If
End Sub

' �{�^���������đS�ẴR���{�{�b�N�X�̑I����ύX
Private Sub SelectionButton_Click()
    For i = 1 To 31
        Me.Controls("CheckBox" & i) = True
    Next
End Sub

' �{�^���������đS�ẴR���{�{�b�N�X�̑I����ύX
Private Sub ReleaseButton_Click()
    For i = 1 To 31
        Me.Controls("CheckBox" & i) = False
    Next
End Sub

' �������I�����ꂽ�珈��
Private Sub NameBox_Change()
    Dim Index As Integer    ' �R���{�{�b�N�X�̔ԍ�
    
    Index = NameBox.ListIndex

    ' ���Ɋ�]�x���I������Ă�����̓`�F�b�N������
    If Index <> -1 Then
            Index = Index + 10
        
        For i = 1 To 31
            If Cells(6 + Index, 8 + i).Value = "��" Then
                Me.Controls("CheckBox" & i) = True
            Else
                Me.Controls("CheckBox" & i) = False
            End If
        Next
    End If
End Sub

' ����{�^��
Private Sub ����_Click()
    Unload Me
End Sub

