VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "scCtlPnl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Function AddRemove()
    Shell "rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5
End Function

Function AddHardWare()
    Shell "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5
End Function

Function TimeDateSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5
End Function

Function RegionalSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5
End Function

Function DisplaySettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5
End Function

Function InternetSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5
End Function

Function KeyboardSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5
End Function

Function MouseSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", 5
End Function

Function ModemSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL modem.cpl", 5
End Function

Function SystemSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5
End Function

Function NetworkSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5
End Function

Function PasswordSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL password.cpl", 5
End Function

Function SoundsSettings()
    Shell "rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5
End Function
