Attribute VB_Name = "Module1"
Option Explicit
Sub �V�[�g�����o�J�n()
    Application.ScreenUpdating = False
    Dim �� As String
    Dim �N�_�t�H���_ As Folder
    Dim �n�� As Date, �I�� As Date
    Dim �I�s As Long
    Dim FSO As New FileSystemObject
    �� = "�N�_�t�H���_���ɂ���S�Ẵt�H���_����Excel�t�@�C���̃V�[�g�������o���܂��B" & vbCrLf & vbCrLf & "�������J�n���Ă�낵���ł����H"
    �� = �� & vbCrLf & "���N�_�t�H���_�ݒ莟��ł͏����Ɏ��Ԃ��|����܂��I"
    If MsgBox(��, vbYesNo) = vbYes Then
        Set �N�_�t�H���_ = FSO.GetFolder(Sheets("�V�[�g������").Cells(2, 1))
        �n�� = Timer
        ���s��.Show vbModeless
        ���s��.Repaint
        Call �V�[�g�����o(�N�_�t�H���_)
        With Sheets("�V�[�g������")
            �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
            .Range("A5:L" & Rows.Count).Borders.LineStyle = False
            .Range("A5:L" & �I�s).Borders.LineStyle = True
        End With
        �I�� = Timer
        Unload ���s��
        MsgBox "���o����" & vbCrLf & "�������ԁF" & �I�� - �n��
    End If
    Application.ScreenUpdating = True
End Sub
Sub �V�[�g�����o(�t�H���_ As Folder)
    Dim �t�@�C�� As File
    Dim ����(1 To 1, 1 To 12)
    Dim �I�s As Long, �� As Long, �� As Long, �V�[�g�� As Long
    Dim �V�[�g As Worksheet
    With Sheets("�V�[�g������")
        For Each �t�@�C�� In �t�H���_.Files
            If �t�@�C��.Name <> ThisWorkbook.Name And InStrRev(�t�@�C��.Name, ".xls") > 0 And InStrRev(�t�@�C��.Name, "~$") = 0 Then
                Workbooks.Open �t�@�C��.Path, UpdateLinks:=0
                ����(1, 1) = �t�H���_.Path
                ����(1, 2) = �t�@�C��.Name
                �V�[�g�� = Workbooks(�t�@�C��.Name).Sheets.Count
                If �V�[�g�� <= 10 Then
                    �� = 3
                    For Each �V�[�g In Workbooks(�t�@�C��.Name).Sheets
                        ����(1, ��) = �V�[�g.Name
                        �� = �� + 1
                    Next
                    Else: ����(1, 3) = "���V�[�g������"
                End If
                Workbooks(�t�@�C��.Name).Close SaveChanges:=False
                �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
                Range(.Cells(�I�s + 1, 1), .Cells(�I�s + 1, 12)) = ����
                Erase ����
            End If
        Next
    End With
    For Each �t�H���_ In �t�H���_.SubFolders
        Call �V�[�g�����o(�t�H���_)
    Next
End Sub
Sub �V�[�g�����o�N���A()
    With Sheets("�V�[�g������")
        Range(.Cells(5, 1), .Cells(Rows.Count, 12)).ClearContents
        Range(.Cells(5, 1), .Cells(Rows.Count, 12)).Borders.LineStyle = False
    End With
End Sub
Sub �V�[�g�A�����o()
    Application.ScreenUpdating = False
    Dim �I�s As Long, �I�� As Long, �s As Long, �� As Long, �Y�s As Long, ��I�s As Long
    Dim �� As String
    Dim �n�� As Date, �I�� As Date
    Dim ��z��, �z��
    With Sheets("�W�J�����ݒ�")
        �I�s = .Cells(Rows.Count, 4).End(xlUp).Row
        If �I�s = 1 Then MsgBox "�u�W�J�����ݒ�v�V�[�g�Ƀf�[�^������܂���": Exit Sub
        �� = "�t�@�C��" & �I�s - 1 & "�����ɂ��ĘA���Łu�W�J���o�v�V�[�g�ɏ��o���܂�" & vbCrLf & vbCrLf & "�������J�n���Ă�낵���ł����H"
        �� = �� & vbCrLf & "���W�J�V�[�g��1��ڂōŉ��s���Z�肵�܂�"
        If MsgBox(��, vbYesNo) = vbNo Then Exit Sub
        �n�� = Timer
        ���s��.Show vbModeless
        ���s��.Repaint
        ReDim ���X�g(2 To �I�s, 1 To 3)
        For �s = 2 To �I�s
            ���X�g(�s, 1) = .Cells(�s, 2) & "\" & .Cells(�s, 3)
            ���X�g(�s, 2) = .Cells(�s, 3)
            ���X�g(�s, 3) = .Cells(�s, 4)
        Next
    End With
    For �Y�s = 2 To �I�s
        Workbooks.Open ���X�g(�Y�s, 1)
        With Workbooks(���X�g(�Y�s, 2))
            With Sheets(���X�g(�Y�s, 3))
                �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
                �I�� = .Cells(1, Columns.Count).End(xlToLeft).Column
                If �I�s = 1 And �I�� = 1 Then
                    ReDim ��z��(1 To 1, 1 To 1)
                    Else: ��z�� = Range(.Cells(1, 1), .Cells(�I�s, �I��))
                End If
            End With
            ReDim �z��(1 To �I�s, 1 To �I�� + 1)
            For �s = 1 To �I�s
                �z��(�s, 1) = .Name
                For �� = 1 To �I��
                    �z��(�s, �� + 1) = ��z��(�s, ��)
                Next
            Next
        End With
        Workbooks(���X�g(�Y�s, 2)).Close
        With ThisWorkbook.Sheets("�W�J���o")
            ��I�s = .Cells(Rows.Count, 1).End(xlUp).Row
            Range(.Cells(��I�s + 1, 1), .Cells(��I�s + �I�s, �I�� + 1)) = �z��
        End With
    Next
    �I�� = Timer
    Unload ���s��
    MsgBox "���o����" & vbCrLf & "�������ԁF" & �I�� - �n��
    Application.ScreenUpdating = True
End Sub
Sub �����ݒ胊�X�g�N���A()
    With Sheets("�W�J�����ݒ�")
        Range(.Cells(2, 2), .Cells(Rows.Count, 4)).ClearContents
    End With
End Sub
Sub �W�J���o�N���A()
    With Sheets("�W�J���o")
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
    End With
End Sub
