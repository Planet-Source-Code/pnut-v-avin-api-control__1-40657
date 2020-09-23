VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScreen"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean


Function CurrentResolution(FormHDC As Long) As String
    Dim D() As DEVMODE
    Dim lNumModes As Long
    Dim Temp As String
    Dim l As Long, lMaxModes As Long
    Dim lBits As Long, lWidth As Long, lHeight As Long
    
    lBits = GetDeviceCaps(FormHDC, 12)
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    lMaxModes = 8
    
    ReDim D(0 To lMaxModes) As DEVMODE
    
    lNumModes = 0
    
    l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    
    Do While l
        If lBits = D(lNumModes).dmBitsPerPel And lWidth = D(lNumModes).dmPelsWidth And lHeight = D(lNumModes).dmPelsHeight Then Temp = D(lNumModes).dmPelsWidth & "x" & D(lNumModes).dmPelsHeight & "x" & D(lNumModes).dmBitsPerPel
        lNumModes = lNumModes + 1
        If lNumModes > lMaxModes Then
            lMaxModes = lMaxModes + 8
            ReDim Preserve D(0 To lMaxModes) As DEVMODE
        End If
        l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Loop

    CurrentResolution = Temp
End Function

Function AllResolutions(Seperator As String, FormHDC As Long) As String
    Dim D() As DEVMODE
    Dim lNumModes As Long
    Dim Temp As String
    Dim l As Long, lMaxModes As Long
    Dim lBits As Long, lWidth As Long, lHeight As Long
    
    If Seperator = "" Then Exit Function
    
    lBits = GetDeviceCaps(FormHDC, 12)
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    lMaxModes = 8
    
    ReDim D(0 To lMaxModes) As DEVMODE
    
    lNumModes = 0
    l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    
    Do While l
        Temp = Temp & D(lNumModes).dmPelsWidth & "x" & D(lNumModes).dmPelsHeight & "x" & D(lNumModes).dmBitsPerPel & Seperator
        lNumModes = lNumModes + 1
        If lNumModes > lMaxModes Then
            lMaxModes = lMaxModes + 8
            ReDim Preserve D(0 To lMaxModes) As DEVMODE
        End If
        l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Loop
    Temp = Left$(Temp, Len(Temp) - Len(Seperator))
    
    AllResolutions = Temp
End Function

Function SetResolution(FormHDC As Long, Wdth As Long, Hght As Long, BPP As Long)
    Dim D() As DEVMODE
    Dim l As Long, Flags As Long, X As Long, lMaxModes As Long
    Dim lBits As Long, lWidth As Long, lHeight As Long
    
    lBits = GetDeviceCaps(FormHDC, 12)
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    lMaxModes = 8
    ReDim D(0 To lMaxModes) As DEVMODE
    lNumModes = 0
    l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Do While l
        lNumModes = lNumModes + 1
        If lNumModes > lMaxModes Then
            lMaxModes = lMaxModes + 8
            ReDim Preserve D(0 To lMaxModes) As DEVMODE
        End If
        l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Loop
    lNumModes = lNumModes - 1
    
    For X = 0 To lMaxModes
        If D(X).dmPelsHeight = Hght And D(X).dmPelsWidth = Wdth And D(X).dmBitsPerPel = BPP Then Exit For
    Next X
    If X > lMaxModes Then Exit Function
    
    D(X).dmFields = &H40000 Or &H80000 Or &H100000
    Flags = &H1
    l = ChangeDisplaySettings(D(X), Flags)
    Select Case l
        Case 1
            l = MsgBox("This change will not take effect until you reboot the system.  Reboot now?", vbYesNo)
            If l = vbYes Then
                Flags = 0
                l = ExitWindowsEx(2, Flags)
            End If
        Case 0
        Case Else
            MsgBox "Error changing resolution! Returned: " & l
    End Select
End Function

