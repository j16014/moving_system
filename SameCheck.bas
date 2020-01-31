Attribute VB_Name = "SameCheck"
Option Explicit

' �������݃`�F�b�N
Public Function exist_check(Work As String, Unit As String) As Boolean
    ' ���݊J���Ă��郏�[�N�u�b�N�E���[�N�V�[�g��S�Ď擾���A�����Ƃ��ēn���ꂽ���O�Ɠ����̃��[�N�u�b�N�E���[�N�V�[�g�����݂��邩�𔻒肷��B�߂�l�Ƃ��ău�[���l��Ԃ��B
    Dim wb As Workbook      ' ���[�N�u�b�N
    Dim ws As Worksheet     ' ���[�N�V�[�g
    Dim SameFlg As Boolean  ' �����t���O
    
    SameFlg = False
    
    Select Case Unit
        Case "wb"
            For Each wb In Workbooks
                If wb.name = Work Then
                    SameFlg = True
                End If
            Next
        Case "ws"
            For Each ws In Sheets
                If ws.name = Work Then
                    SameFlg = True
                End If
            Next
    End Select

    exist_check = SameFlg
End Function
