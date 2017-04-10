'--------------------------------------------------
'st.vbs
'--------------------------------------------------
'ModuleName:    SetShortcutLink Module
'FileName:      SetShortcutLink.vbs
'--------------------------------------------------
'Version:       2015/09/07
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'��Include st.vbs
'--------------------------------------------------
Sub Include(ByVal FileName)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim Stream: Set Stream = fso.OpenTextFile( _
        fso.BuildPath( _
            fso.GetParentFolderName(WScript.ScriptFullName), _
            FileName) , 1)
    Call ExecuteGlobal(Stream.ReadAll())
    Call Stream.Close
End Sub
'--------------------------------------------------
Call Include("..\Lib\st.vbs")
'--------------------------------------------------

Call Main

Sub Main
Do

    Dim ProgramFileName: ProgramFileName = _
        LastStrLastDelim(fso.GetBaseName(ScriptFilePath), "_")

    If ProgramFileName = "" Then
        Call MsgBox(StringCombine(vbCrLf, Array( _
            "�G���[:�t�@�C����������������܂���", _
            fso.GetBaseName(ScriptFilePath))))
        Exit Do
    End If

    Dim TargetSourceFileName: TargetSourceFileName = _
        ProgramFileName + ".vbs"

    Dim ShortcutLinkFileName: ShortcutLinkFileName = _
        ProgramFileName + ".lnk"

    Dim StartMenuGroupName: StartMenuGroupName = _
        fso.BuildPath(ProgramFileName, ShortcutLinkFileName)

    '--------------------------------------------------

    Dim ShortcutLinkFilePathList: ShortcutLinkFilePathList = ""
    Dim DeleteFolderPathList: DeleteFolderPathList = ""

    ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
        fso.BuildPath(SendToFolderPath, ShortcutLinkFileName)))

    ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
        fso.BuildPath(DesktopFolderPath, ShortcutLinkFileName)))

    ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
        fso.BuildPath(StartMenuProgramsFolderPath, StartMenuGroupName)))
    DeleteFolderPathList = StringCombine(vbCrLf, Array(DeleteFolderPathList, _
        fso.BuildPath(StartMenuProgramsFolderPath, ProgramFileName)))
    '�폜�\���̂���t�H���_��DeleteFolderPathList�ɓo�^����B
    '�X�^�[�g���j���[�̃v���W�F�N�g�O���[�v�t�H���_�͍폜�\��������B

    ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
        fso.BuildPath(StartMenuFolderPath, ShortcutLinkFileName)))

    ShortcutLinkFilePathList = StringCombine(vbCrLf, Array(ShortcutLinkFilePathList, _
        fso.BuildPath(StartUpFolderPath, ShortcutLinkFileName)))
    '--------------------------------------------------

    Dim TargetSourceFilePath: TargetSourceFilePath = _
        fso.BuildPath( _
            fso.GetParentFolderName(ScriptFolderPath), _
            TargetSourceFileName)

    If fso.FileExists(TargetSourceFilePath) = False Then
        Call MsgBox(StringCombine(vbCrLf, Array( _
            "�t�@�C�������݂��܂���B", _
            TargetSourceFilePath )))
        Exit Do
    End If

    Dim I

    Dim ShortcutLinkFilePathArray
    ShortcutLinkFilePathArray = Split(ShortcutLinkFilePathList, vbCrLf)
    For I = 0 To ArrayCount(ShortcutLinkFilePathArray) - 1
    Do
        Dim ShortcutLinkFilePath: ShortcutLinkFilePath = _
            ShortcutLinkFilePathArray(I)
        Dim ShortcutLinkFileParentFolderPath: ShortcutLinkFileParentFolderPath = _
            fso.GetParentFolderName(ShortcutLinkFilePath)

        Select Case MsgBox(StringCombine(vbCrLf, Array( _
            "�V���[�g�J�b�g�t�@�C�����쐬���܂����H", _
            fso.GetFileName(ShortcutLinkFileParentFolderPath), _
            "�͂�=�쐬����", _
            "������=�폜����")), vbYesNo)
        Case vbYes
            Call ForceCreateFolder(ShortcutLinkFileParentFolderPath)

            Call CreateShortcutFile( _
                ShortcutLinkFilePath, _
                TargetSourceFilePath, _
                ScriptProgramFilePath, 2, _
                "")
            Call ShellFileOpen( _
                ShortcutLinkFileParentFolderPath, _
                vbNormalFocus, True)
        Case vbNo
            If fso.FileExists(ShortcutLinkFilePath) Then
                Call fso.DeleteFile(ShortcutLinkFilePath)
                Call ShellFileOpen( _
                    ShortcutLinkFileParentFolderPath, _
                    vbNormalFocus, True)
            End If
        End Select

    Loop While False
    Next

    '�X�^�[�g���j���[�v���O�����O���[�v�t�H���_�̍폜����
    Dim DeleteFolderPathArray
    DeleteFolderPathArray = Split(DeleteFolderPathList, vbCrLf)
    For I = 0 To ArrayCount(DeleteFolderPathArray) - 1
    Do
        Dim DeleteFolderPath: DeleteFolderPath = _
            DeleteFolderPathArray(I)
        Dim DeleteFolderParentFolderPath: DeleteFolderParentFolderPath = _
            fso.GetParentFolderName(DeleteFolderPath)

        If fso.FolderExists(DeleteFolderPath) Then
            If FolderHasSubItem(DeleteFolderPath) = False Then
                Call fso.DeleteFolder(DeleteFolderPath)
                'Call ShellFileOpen( _
                '    DeleteFolderParentFolderPath, _
                '    vbNormalFocus, True)
            End If
        End If

    Loop While False
    Next


    Call MsgBox( _
        "Finish " + WScript.ScriptName)

Loop While False
End Sub

