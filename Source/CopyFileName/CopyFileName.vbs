'--------------------------------------------------
'CopyFileName
'ModuleName:    CopyFileName.vbs
'--------------------------------------------------
'Version:       2015/11/05
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
Call Include(".\Common\CopyFilePathNameCommon.vbs")
'--------------------------------------------------

Call Main(Enum_CopyFilePathType_Name)
