Option Explicit

'--------------------------------------------------
'��Include st.vbs
'--------------------------------------------------
Sub Include(ByVal FileName)
    Dim fso: Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
    Dim Stream: Set Stream = fso.OpenTextFile( _
        fso.GetParentFolderName(WScript.ScriptFullName) _
        + "\" + FileName, 1)
    Call ExecuteGlobal(Stream.ReadAll())
    Call Stream.Close
End Sub
'--------------------------------------------------
Call Include(".\Lib\st.vbs")
'--------------------------------------------------

'------------------------------
'�����C������
'------------------------------
Call Main

Sub Main
    Dim MessageText: MessageText = ""

    Dim IniFilePath: IniFilePath = _
        PathCombine(Array(ScriptFolderPath, "SupportTool.ini"))

    Dim IniFile: Set IniFile = New IniFile
    Call IniFile.Initialize(IniFilePath)

    '--------------------
    '�E�ݒ�Ǎ�
    '--------------------
    Dim ProjectName: ProjectName = _
        IniFile.ReadString("Common", "ProjectName", "")
    If ProjectName = "" Then
        WScript.Echo _
            "�ݒ肪�ǂݎ��Ă��܂���"
        Exit Sub
    End If

    Dim InstallParentFolderPath: InstallParentFolderPath = _
        IniFile.ReadString("ReleaseInstall", "InstallParentFolderPath", "")
    If InstallParentFolderPath = "" Then
        WScript.Echo _
            "�ݒ肪�ǂݎ��Ă��܂���"
        Exit Sub
    End If

    Dim InstallFolderName: InstallFolderName = _
        IniFile.ReadString("ReleaseInstall", "InstallFolderName", "")

    Dim IgnoreFiles: IgnoreFiles = _
        IniFile.ReadString("ReleaseInstall", "InstallIgnoreFiles", "")

    Dim OverWriteIgnoreFiles: OverWriteIgnoreFiles = _
        IniFile.ReadString("ReleaseInstall", "InstallOverWriteIgnoreFiles", "")
    '--------------------

    Dim NowValue: NowValue = Now
    Dim ReleaseFolderPath: ReleaseFolderPath = _
        PathCombine(Array( _
            "..\..\Release", _
            "Recent", _
            ProjectName))
    ReleaseFolderPath = _
        AbsoluteFilePath(ScriptFolderPath, ReleaseFolderPath)

    Dim InstallFolderPath: InstallFolderPath = _
        PathCombine(Array( _
            InstallParentFolderPath, _
            IIF(InstallFolderName="", ProjectName, InstallFolderName)))

    If not fso.FolderExists(ReleaseFolderPath) Then
        WScript.Echo _
            "�R�s�[���t�H���_��������܂���" + vbCrLF + _
            ReleaseFolderPath
        Exit Sub
    End If

    If not fso.FolderExists(InstallParentFolderPath) Then
        WScript.Echo _
            "�C���X�g�[����e�t�H���_��������܂���" + vbCrLF + _
            InstallParentFolderPath
        Exit Sub
    End If

    Call CopyFolderIgnorePath( _
        ReleaseFolderPath, InstallFolderPath, _
        IgnoreFiles, OverWriteIgnoreFiles)

    MessageText = MessageText + _
        fso.GetFileName(InstallFolderPath) + vbCrLf

    Dim MessageResult: MessageResult = _
        MsgBox(StringCombine(vbCrLf, Array( _
            "�t�H���_���J���܂����H", _
            "Finish " + WScript.ScriptName, _
            "----------", _
            Trim(MessageText) )), vbYesNo)
    If MessageResult = vbYes Then
        Call ShellFileOpen(InstallFolderPath, vbNormalFocus, True)
    End If
End Sub

