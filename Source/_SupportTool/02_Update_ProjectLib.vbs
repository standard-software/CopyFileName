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
    Dim LibrarySourceFilePathList: LibrarySourceFilePathList = ""
    Dim LibraryDestFilePathList: LibraryDestFilePathList = ""

    'LibrarySourcePath01/02/03...�Ƃ���Ini�t�@�C�����ڂ̓ǂݎ��
    Dim I: I = 1
    Do While IniFile.SectionIdentExists( _
        "Update_ProjectLib", _
        "LibrarySourceFilePath" + LongToStrDigitZero(I, 2))

        If IniFile.SectionIdentExists( _
            "Update_ProjectLib", _
            "LibraryDestFilePath" + LongToStrDigitZero(I, 2)) = False Then
            WScript.Echo StringCombine(vbCrLf, Array( _
                "�ݒ肪����������܂���B", _
                "LibrarySourceFilePath - LibraryDestFilePath:" + CStr(I)))
            Exit Sub
        End If

        LibrarySourceFilePathList = _
            StringCombine(vbCrLf, Array(LibrarySourceFilePathList, _
            IniFile.ReadString( _
                "Update_ProjectLib", _
                "LibrarySourceFilePath"  + LongToStrDigitZero(I, 2), "")))

        LibraryDestFilePathList = _
            StringCombine(vbCrLf, Array(LibraryDestFilePathList, _
            IniFile.ReadString( _
                "Update_ProjectLib", _
                "LibraryDestFilePath"  + LongToStrDigitZero(I, 2), "")))

        I = I + 1
    Loop

    If LibrarySourceFilePathList = "" Then
        WScript.Echo _
            "�ݒ肪�ǂݎ��Ă��܂���"
        Exit Sub
    End If

    Dim ProjectName: ProjectName = _
        IniFile.ReadString("Common", "ProjectName", "")
    If ProjectName = "" Then
        WScript.Echo _
            "�ݒ肪�ǂݎ��Ă��܂���"
        Exit Sub
    End If

    'Dim LibraryDestFolderPath: LibraryDestFolderPath = _
    '    IniFile.ReadString("Update_ProjectLib", "LibraryDestFolderPath", "")
    'If LibraryDestFolderPath = "" Then
    '    WScript.Echo _
    '        "�ݒ肪�ǂݎ��Ă��܂���"
    '    Exit Sub
    'End If
    '--------------------

    'Dim DestFolderPath: DestFolderPath = _
    '    AbsolutePath(ScriptFolderPath, LibraryDestFolderPath)
    'If not fso.FolderExists(DestFolderPath) Then
    '    WScript.Echo _
    '        "�R�s�[��t�H���_��������܂���" + vbCrLF + _
    '        fso.GetParentFolderName(DestFolderPath)
    '    Exit Sub
    'End If

    Dim DestFilePaths: DestFilePaths = Split(LibraryDestFilePathList, vbCrLf)


    Dim SourceFilePaths: SourceFilePaths = Split(LibrarySourceFilePathList, vbCrLf)

    If ArrayCount(SourceFilePaths) <> ArrayCount(DestFilePaths) Then
        WScript.Echo StringCombine(vbCrLf, Array( _
            "�ݒ肪����������܂���B", _
            "SourceFilePaths.Count <> DestFilePaths.Count"))
        Exit Sub
    End If

    Dim SourceFilePath
    Dim DestFilePath
    For I = 0 To ArrayCount(SourceFilePaths) - 1
        SourceFilePath = SourceFilePaths(I)
        DestFilePath = DestFilePaths(I)

        SourceFilePath = _
            AbsolutePath(ScriptFolderPath, SourceFilePath)
        DestFilePath = _
            AbsolutePath(ScriptFolderPath, DestFilePath)

        If not fso.FileExists(SourceFilePath) Then
            WScript.Echo StringCombine(vbCrLf, Array( _
                "�R�s�[���t�@�C����������܂���", _
                SourceFilePath))
            Exit Sub
        End If

        Call fso.CopyFile(SourceFilePath, DestFilePath, True)
        MessageText = StringCombine(vbCrLf, Array(MessageText,  _
            fso.GetFileName(SourceFilePath) + ">> " + _
            fso.GetFileName(fso.GetParentFolderName( DestFilePath ) )))
    Next

    WScript.Echo StringCombine(vbCrLf, Array( _
        "Finish " + WScript.ScriptName, _
        "----------", _
        Trim(MessageText) ))
End Sub