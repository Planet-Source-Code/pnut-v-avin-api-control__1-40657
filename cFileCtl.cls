VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileCtl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Function EmptyRecycle(hWnd As Long)
    SHEmptyRecycleBin hWnd, "", &H2
End Function

Function FindFiles()
    keybd_event &H5B, 0, 0, 0
    keybd_event 70, 0, 0, 0
    keybd_event &H5B, 0, &H2, 0
End Function

Function Copy(Filename As String, Destination As String)
    CopyFile Filename, Destination, 1
End Function

Function Delete(Filename As String)
    DeleteFile Filename
End Function

Function Move(Filename As String, Destination As String)
    MoveFile Filename, Destination
End Function

Function DoesPathExist(Filepath As String) As Boolean
    DoesPathExist = PathFileExists(Filepath)
End Function

