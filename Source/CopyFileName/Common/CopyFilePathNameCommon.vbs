Option Explicit

Const Enum_CopyFilePathType_FullPath = 1
Const Enum_CopyFilePathType_Name = 2

Sub Main(ByVal CopyFilePathType)
Do
    Dim ArgsArray
    ArgsArray = ArgsToArray

    '����m�F�p�R�[�h
    'ArgsArray = Array( _
    '    fso.BuildPath(ScriptFolderPath, "CopyFileName.vbs"), _
    '    fso.BuildPath(ScriptFolderPath, "CopyFilePath.vbs"))

    'MsgBox ArrayToString(ArgsArray, " ")
    'Exit Sub

    If ArrayCount(ArgsArray) = 0 Then
        Call WScript.Echo("Args.Count = 0")
        Exit Sub
    End If

    Dim FileArrayList
    Set FileArrayList = CreateObject("System.Collections.ArrayList")
    
    Dim I

    '�V���[�g�J�b�g�t�@�C�����܂܂�Ă��邩�ǂ���������
    Dim ShortcutLinkFlag: ShortcutLinkFlag = False
    For I = 0 To ArrayCount(ArgsArray) - 1
        If fso.FileExists(ArgsArray(I)) Then
            If IsShortcutLinkFile(ArgsArray(I)) Then
                ShortcutLinkFlag = True
                Exit For
            End If 
        End If
    Next

    '�V���[�g�J�b�g�t�@�C����W�J���邩�ǂ������߂�
    Dim ShortcutLinkSourceFlag: ShortcutLinkSourceFlag = False
    If ShortcutLinkFlag Then
        If vbYes = MsgBox( _
            "�V���[�g�J�b�g�t�@�C���̃����N��p�X���擾���܂����H", _
            vbYesNo) Then
            'Message:Get Path ShortcutFile Link Source ?
            ShortcutLinkSourceFlag = True
        End If
    End If

    For I = 0 To ArrayCount(ArgsArray) - 1
        If fso.FileExists(ArgsArray(I)) Then
            If ShortcutLinkSourceFlag _
            And IsShortcutLinkFile(ArgsArray(I)) Then
                Call FileArrayList.Add( _
                    PathConvert(CopyFilePathType, _
                    ShortcutFileLinkPath(ArgsArray(I))))
            Else
                Call FileArrayList.Add( _
                    PathConvert(CopyFilePathType, ArgsArray(I)))
            End If
        ElseIf fso.FolderExists(ArgsArray(I)) Then
            Call FileArrayList.Add( _
                PathConvert(ArgsArray(I)))
        End If
    Next

    '�\�[�g
    Call FileArrayList.Sort

    Dim CopyText: CopyText = ""
    For I = 0  To FileArrayList.Count - 1
        CopyText = CopyText + FileArrayList(I) + vbCrLf
    Next

    Call SetClipboardText(CopyText)
    
    Call WScript.Echo( _
        "�N���b�v�{�[�h�Ƀt�@�C�������R�s�[���܂����B" _
         + vbCrLf + CopyText)
        'Message:Copy Text To Clipboard.

Loop While False
End Sub

Private Function PathConvert(ByVal CopyFilePathType, ByVal FilePath)
    Dim Result: Result = ""
    Select Case CopyFilePathType
    Case Enum_CopyFilePathType_FullPath:
        Result = FilePath
    Case Enum_CopyFilePathType_Name:
        Result = fso.GetFileName(FilePath)
    End Select
    PathConvert = Result
End Function

