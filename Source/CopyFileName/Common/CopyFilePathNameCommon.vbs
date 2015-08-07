Option Explicit

Const Enum_CopyFilePathType_FullPath = 1
Const Enum_CopyFilePathType_Name = 2

Sub Main(ByVal CopyFilePathType)
Do
    Dim Args: Set Args = WScript.Arguments
    
    If Args.Count = 0 Then
        Call WScript.Echo("Args.Count = 0")
        Exit Sub
    End If
    
    Dim FileArrayList
    Set FileArrayList = CreateObject("System.Collections.ArrayList")
    
    Dim I

    'ショートカットファイルが含まれているかどうかを検索
    Dim ShortcutLinkFlag: ShortcutLinkFlag = False
    For I = 0 To Args.Count - 1
        If fso.FileExists(Args(I)) Then
            If IsShortcutLinkFile(Args(I)) Then
                ShortcutLinkFlag = True
                Exit For
            End If 
        End If
    Next

    'ショートカットファイルを展開するかどうか決める
    Dim ShortcutLinkSourceFlag: ShortcutLinkSourceFlag = False
    If ShortcutLinkFlag Then
        If vbYes = MsgBox( _
            "ショートカットファイルのリンク先パスを取得しますか？", _
            vbYesNo) Then
            'Message:Get Path ShortcutFile Link Source ?
            ShortcutLinkSourceFlag = True
        End If
    End If

    For I = 0 To Args.Count - 1
        If fso.FileExists(Args(I)) Then
            If ShortcutLinkSourceFlag _
            And IsShortcutLinkFile(Args(I)) Then
                Call FileArrayList.Add( _
                    PathConvert(CopyFilePathType, _
                    ShortcutFileLinkPath(Args(I))))
            Else
                Call FileArrayList.Add( _
                    PathConvert(CopyFilePathType, Args(I)))
            End If
        ElseIf fso.FolderExists(Args(I)) Then
            Call FileArrayList.Add( _
                PathConvert(Args(I)))
        End If
    Next

    'ソート
    Call FileArrayList.Sort

    Dim CopyText: CopyText = ""
    For I = 0  To FileArrayList.Count - 1
        CopyText = CopyText + FileArrayList(I) + vbCrLf
    Next

    Call SetClipboardText(CopyText)
    
    Call WScript.Echo( _
        "クリップボードにファイル名をコピーしました。" _
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

