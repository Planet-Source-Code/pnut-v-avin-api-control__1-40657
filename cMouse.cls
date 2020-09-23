VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMouse"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim SB As Integer

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function SwapMouseButton Lib "user32.dll" (ByVal bSwap As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

Function GetCursorX() As Long
    Dim MousePos As POINTAPI
    
    GetCursorPos MousePos
    
    GetCursorX = MousePos.X
End Function

Function GetCursorY() As Long
    Dim MousePos As POINTAPI
    
    GetCursorPos MousePos
    
    GetCursorY = MousePos.Y
End Function

Function FlipMouseButtons()
    If SB = 0 Then SB = 1 Else: SB = 0
    SwapMouseButton SB
End Function

Function CursorShow()
    ShowCursor True
End Function

Function CursorHide()
    ShowCursor False
End Function

Function TrapCursor(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    Dim tReturn As RECT
    
    With tReturn
        .Left = X1: .Top = Y1: .Right = X2: .Bottom = Y2
    End With
    
    ClipCursor tReturn
End Function

Function ReleaseCursor()
    Dim tReturn As RECT
    
    With tReturn
        .Left = 0: .Top = 0: .Right = Screen.Width / Screen.TwipsPerPixelX: .Bottom = Screen.Height / Screen.TwipsPerPixelY
    End With
    
    ClipCursor tReturn
End Function
