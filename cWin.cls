VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWin"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public ControlPanel As New scCtlPnl

Function ShutDown()
    ExitWindowsEx 1, 0&
End Function

Function Restart()
    ExitWindowsEx 2, 0&
End Function

Function LogOff()
    ExitWindowsEx 0, 0&
End Function

Function TaskBarHide()
    Dim tReturn
    
    tReturn = FindWindow("Shell_traywnd", "")
    SetWindowPos tReturn, 0, 0, 0, 0, 0, &H80
End Function

Function TaskBarShow()
    Dim tReturn As Long

    tReturn = FindWindow("Shell_traywnd", "")
    SetWindowPos tReturn, 0, 0, 0, 0, 0, &H40
End Function

Function ScreenSaverOn(FormHWND As Long)
    SendMessage FormHWND, &H112&, &HF140&, 0&
End Function

Function ScreenSaverOff()
    SendKeys Chr$(27), True
End Function

Function DesktopIconsShow()
    Dim hWnd As Long

    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 5
End Function

Function DesktopIconsHide()
    Dim hWnd As Long

    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 0
End Function

' False for on, True for off
Function TurnOffCtlAltDel(OnOff As Boolean)
    SystemParametersInfo 97, OnOff, CStr(1), 0
End Function

Function MinimizeAll()
    keybd_event &H5B, 0, 0, 0
    keybd_event 77, 0, 0, 0
    keybd_event &H5B, 0, &H2, 0
End Function

Function OpenExplore()
    keybd_event &H5B, 0, 0, 0
    keybd_event 69, 0, 0, 0
    keybd_event &H5B, 0, &H2, 0
End Function

Function FormOnTop(FormHWND As Long, Left As Long, Top As Long, Width As Long, Height As Long)
    SetWindowPos FormHWND, -1, Left, Top, Width, Height, &H10 Or &H40
End Function

Function GoToSleep(LenInMilliseconds As Long)
    Sleep LenInMilliseconds
End Function

Function ShutDownDIALOG()
    SHShutDownDialog 0
End Function

Function EnableObject(Object As Object)
    EnableWindow Object.hWnd, True
End Function

Function DisableObject(Object As Object)
    EnableWindow Object.hWnd, False
End Function

Function AddFiletoDocsMenu(Filename As String)
    SHAddToRecentDocs 2, Filename
End Function

Function ClearDocsMenu()
    SHAddToRecentDocs 2, vbNullString
End Function
          
Function GetNameofUser() As String
    Dim Buff As String
    Dim Sze As Long
    
    Buff = Space$(255)
    Sze = 255
    
    GetUserName Buff, Sze
    GetNameofUser = Left$(Buff, Sze)
End Function

Function GetNameofComputer() As String
    Dim Buff As String
    Dim Sze As Long
    
    Buff = Space$(255)
    Sze = 255
    
    GetComputerName Buff, Sze
    GetNameofComputer = Left$(Buff, Sze)
End Function

Function GetTimeinWindows() As String
    GetTimeinWindows = CStr(Round((GetTickCount / 1000 / 60))) & " Minutes in Windows"
End Function

Function RefreshWindows()
    Dim tReturn As Long
    Dim hWnd As Long
    
    rtn = FindWindow("Shell_traywnd", "")
    SetWindowPos rtn, 0, 0, 0, 0, 0, &H40
    SystemParametersInfo 97, False, CStr(1), 0
    
    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 5
End Function

Function HideTime()
    Dim P As Long, c As Long, a As Long
    P = FindWindow("Shell_TrayWnd", vbNullString)
    c = FindWindowEx(P&, 0&, "TrayNotifyWnd", vbNullString)
    a = FindWindowEx(c&, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(a&, 0)
End Function

Function ShowTime()
    Dim P As Long, c As Long, a As Long
    P = FindWindow("Shell_TrayWnd", vbNullString)
    c = FindWindowEx(P, 0&, "TrayNotifyWnd", vbNullString)
    a = FindWindowEx(c, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(a, 5)
End Function

Public Sub ShowStart()
    Dim Taskbar As Long, StartButton As Long
    
    Taskbar& = FindWindow("Shell_TrayWnd", vbNullString)
    StartButton& = FindWindowEx(Taskbar&, 0&, "Button", vbNullString)
    
    ShowWindow StartButton&, 5
End Sub

Public Sub HideStart()
    Dim Taskbar As Long, StartButton As Long
    
    Taskbar& = FindWindow("Shell_TrayWnd", vbNullString)
    StartButton& = FindWindowEx(Taskbar&, 0&, "Button", vbNullString)
    
    ShowWindow StartButton&, 0
End Sub
