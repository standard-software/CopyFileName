'--------------------------------------------------
'st.vbs
'--------------------------------------------------
'ModuleName:    SetShortcutLink Module
'FileName:      SetShortcutLink.vbs
'--------------------------------------------------
'Version:       2015/08/24
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
Call Include(".\Lib\st.vbs")
'--------------------------------------------------

'--------------------------------------------------
Dim TargetSourceFileName: TargetSourceFileName = _
    "CopyFileName.vbs"

Dim ShortcutLinkFileName: ShortcutLinkFileName = _
    "CopyFileName.lnk"
'StartMenu\Programs�̒��̏ꍇ�� �v���O�����O���[�v���w�肷�邽�߂�
'"Project01\Project01.lnk" �Ǝw�肷��̂��悢�ł��傤�B

Dim SpecialFolderPath
SpecialFolderPath = SendToFolderPath
'SpecialFolderPath = DesktopFolderPath
'SpecialFolderPath = StartMenuFolderPath
'SpecialFolderPath = StartMenuProgramsFolderPath
'SpecialFolderPath = StartUpFolderPath

Dim ShortcutLinkFilePath: ShortcutLinkFilePath = _
    fso.BuildPath( _
        SpecialFolderPath, _
        ShortcutLinkFileName)

Call Main

Sub Main
Do
    Dim TargetSourceFilePath: TargetSourceFilePath = _
        fso.BuildPath( _
            fso.GetParentFolderName(WScript.ScriptFullName), _
            TargetSourceFileName)

    If fso.FileExists(TargetSourceFilePath) = False Then
        Call MsgBox(StringCombine(vbCrLf, Array( _
            "�t�@�C�������݂��܂���B", _
            TargetSourceFilePath )))
        Exit Do
    End If

    If fso.folderExists(SpecialFolderPath) = False Then
        Call MsgBox(StringCombine(vbCrLf, Array( _
            "�t�H���_�����݂��܂���B", _
            SpecialFolderPath )))
        Exit Do
    End If

    'Call MsgBox(ShortcutLinkFilePath)

    If fso.FileExists(ShortcutLinkFilePath) Then
        fso.DeleteFile(ShortcutLinkFilePath)
    End If

    Call CreateShortcutFile( _
        ShortcutLinkFilePath, _
        TargetSourceFilePath, _
        ScriptProgramFilePath, 2, _
        "")

     Call ShellFileOpen( _
        fso.GetParentFolderName(ShortcutLinkFilePath), _
        vbNormalFocus, True)

Loop While False
End Sub


