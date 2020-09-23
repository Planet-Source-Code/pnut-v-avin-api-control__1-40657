VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHardware"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Const REG_DWORD = 4
Private Const HKEY_DYN_DATA = &H80000006
Private CPULoad As Boolean

Function OpenCDROM()
    Dim tReturn As Long
    
    mciSendString "set CDAudio door open", tReturn, 127, 0
End Function

Function CloseCDROM()
    Dim tReturn As Long
    
    mciSendString "set CDAudio door closed", tReturn, 127, 0
End Function

Function GetCPUUsage() As Long
    Dim Data As Long, Typ As Long, size As Long
    Dim hKey As Long
    Dim hRet As Long
    
    If CPULoad = False Then InitCPU: CPULoad = True

    hRet = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
    hRet = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0&, REG_DWORD, Data, 4)
    GetCPUUsage = Data
    hRet = RegCloseKey(hKey)
End Function

Private Function InitCPU()
    Dim Data As Long, Typ As Long, size As Long
    Dim hKey As Long, hRet As Long

    hRet = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
    hRet = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, REG_DWORD, Data, 4)
    hRet = RegCloseKey(hKey)
End Function
