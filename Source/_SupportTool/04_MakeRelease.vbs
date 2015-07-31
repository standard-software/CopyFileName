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

    '------------------------------
    '�E�ݒ�Ǎ�
    '------------------------------

    Dim ProjectName: ProjectName = _
        IniFile.ReadString("Common", "ProjectName", "")
    If ProjectName = "" Then
        WScript.Echo _
            "�ݒ肪�ǂݎ��Ă��܂���"
        Exit Sub
    End If

    Dim IgnoreFileFolderName: IgnoreFileFolderName = _
        IniFile.ReadString("MakeRelease", "ReleaseIgnoreFileFolderName", "")

    Dim IncludeFileFolderPath: IncludeFileFolderPath = _
        IniFile.ReadString("MakeRelease", "ReleaseIncludeFileFolderPath", "")
    If IncludeFileFolderPath = "" Then
        WScript.Echo _
            "�ݒ肪�ǂݎ��Ă��܂���"
        Exit Sub
    End If

    Dim ScriptEncoderExePath: ScriptEncoderExePath = _
        IniFile.ReadString("MakeRelease", "ScriptEncoderExePath", "")

    Dim ScriptEncodeTargets: ScriptEncodeTargets = _
        IniFile.ReadString("MakeRelease", "ScriptEncodeTargets", "")
    '------------------------------

    Dim NowValue: NowValue = Now
    Dim ReleaseFolderPath: ReleaseFolderPath = _
        PathCombine(Array( _
            "..\..\Release", _
            "Recent", _
            ProjectName))
    ReleaseFolderPath = _
        AbsoluteFilePath(ScriptFolderPath, ReleaseFolderPath)

    Dim SourceFolderPath: SourceFolderPath = _
        "..\..\Source\" + _
        ProjectName
    SourceFolderPath = _
        AbsoluteFilePath(ScriptFolderPath, SourceFolderPath)

    If not fso.FolderExists(SourceFolderPath) Then
        WScript.Echo _
            "�R�s�[���t�H���_��������܂���" + vbCrLF + _
            SourceFolderPath
        Exit Sub
    End If

    '�t�H���_�Đ����R�s�[
'    Call ReCreateCopyFolder(SourceFolderPath, ReleaseFolderPath)
    Call RecreateFolder(fso.GetParentFolderName(ReleaseFolderPath))
    Call CopyFolderIgnorePath( _
        SourceFolderPath, ReleaseFolderPath, _
        IgnoreFileFolderName, "")

    MessageText = MessageText + _
        fso.GetFileName(SourceFolderPath) + vbCrLf

    '�o�[�W�������t�@�C����Readme.txt�̎w��Ȃ�
    Dim IncludeFileFolderArray: IncludeFileFolderArray = _
        Split(IncludeFileFolderPath, ",")
    Dim I
    For I = 0 To ArrayCount(IncludeFileFolderArray) - 1
        IncludeFileFolderArray(I) = _
            AbsoluteFilePath(ScriptFolderPath, IncludeFileFolderArray(I))
        If fso.FileExists(IncludeFileFolderArray(I)) Then
            Call ForceCreateFolder(ReleaseFolderPath)
            Call fso.CopyFile( _
                IncludeFileFolderArray(I), _
                    IncludeLastPathDelim(ReleaseFolderPath) + _
                    fso.GetFileName(IncludeFileFolderArray(I)))
        ElseIf fso.FolderExists(IncludeFileFolderArray(I)) Then
            Call ForceCreateFolder(ReleaseFolderPath)
            Call fso.CopyFolder( _
                IncludeFileFolderArray(I), _
                    IncludeLastPathDelim(ReleaseFolderPath) + _
                    fso.GetFileName(IncludeFileFolderArray(I)))
        Else
            MessageText = StringCombine(vbCrLf, Array( _
                MessageText, _
                "Warning:Include File/Folder not found."))
        End If
    Next

    '�X�N���v�g�G���R�[�h
    Dim ScriptEncodeTargetArray: ScriptEncodeTargetArray = _
        Split(ScriptEncodeTargets, ",")
    If (1 <= ArrayCount(ScriptEncodeTargetArray)) Then
        If fso.FileExists(ScriptEncoderExePath) = False Then
            MessageText = StringCombine(vbCrLf, Array(MessageText, _
                "Warning:ScriptEncoder not found."))
        Else
            For I = 0 To ArrayCount(ScriptEncodeTargetArray) - 1
            Do
                ScriptEncodeTargetArray(I) = AbsoluteFilePath( _ 
                    ScriptFolderPath, ScriptEncodeTargetArray(I))

                '�R�s�[��̃����[�X�t�H���_�ŏ������s��
                ScriptEncodeTargetArray(I) = Replace(ScriptEncodeTargetArray(I), _
                    SourceFolderPath, ReleaseFolderPath)

                If fso.FileExists(ScriptEncodeTargetArray(I)) = False Then
                    MessageText = StringCombine(vbCrLf, Array(MessageText, _
                        "Warning:ScriptTarget not found."))
                    Exit Do
                End If

                Call IncludeExpanded(ScriptEncodeTargetArray(I), ScriptEncodeTargetArray(I))
                Call EncodeVBScriptFile( _
                    ScriptEncoderExePath, _
                    ScriptEncodeTargetArray(I), _
                    ChangeFileExt(ScriptEncodeTargetArray(I), ".vbe"))
            Loop While False
            Next
            If MessageText = "" Then
                For I = 0 To ArrayCount(ScriptEncodeTargetArray) - 1
                    Call ForceDeleteFile(ScriptEncodeTargetArray(I))
                    Call ForceDeleteFolder( _
                        PathCombine(Array( _
                        fso.GetParentFolderName(ScriptEncodeTargetArray(I)), _
                        "Lib")) )
                    'vbs�t�@�C����Lib�t�H���_�̍폜
                Next
            End If
        End If
    End If

    Call WScript.Echo(StringCombine(vbCrLf, Array( _
        "Finish " + WScript.ScriptName, _
        "----------", _
        MessageText)))

End Sub

