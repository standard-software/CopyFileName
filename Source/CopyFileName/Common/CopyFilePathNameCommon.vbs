Option Explicit

Const Enum_CopyFilePath_FullPath = 1
Const Enum_CopyFilePath_Name = 2

Sub Main(ByVal CopyFilePath)
Do
    Dim Args: Set Args = WScript.Arguments
    
    If Args.Count = 0 Then
        Call WScript.Echo("Args.Count = 0")
        Exit Sub
    End If
    
    Dim ArrayList1
    Set ArrayList1 = CreateObject("System.Collections.ArrayList")
    
    Dim I
    For I = 0 To Args.Count - 1
        If fso.FileExists(Args(I)) Or fso.FolderExists(Args(I)) Then
            Select Case CopyFilePath
            Case Enum_CopyFilePath_FullPath:
                Call ArrayList1.Add(Args(I))
            Case Enum_CopyFilePath_Name:
                Call ArrayList1.Add(fso.GetFileName(Args(I)))
            End Select
        End If
    Next

    'É\Å[Ég
    Call ArrayList1.Sort

    Dim CopyText: CopyText = ""
    For I = 0  To ArrayList1.Count - 1
        CopyText = CopyText + ArrayList1(I) + vbCrLf
    Next

    Call SetClipboardText(CopyText)
    
    Call WScript.Echo("Copy Text To Clipboard." + vbCrLf + CopyText)
Loop While False
End Sub


