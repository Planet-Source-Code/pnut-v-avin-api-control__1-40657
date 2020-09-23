VERSION 5.00
Begin VB.UserControl API 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   InvisibleAtRuntime=   -1  'True
   Picture         =   "API.ctx":0000
   ScaleHeight     =   88
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "API.ctx":030A
End
Attribute VB_Name = "API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public ScreenControl As New cScreen
Public FileControl As New cFileCtl
Public Hardware As New cHardware
Public Windows As New cWin
Public Internet As New cInternet
Public Mouse As New cMouse
Public Media As New cMedia

Private Sub UserControl_Resize()
    UserControl.Width = 32 * Screen.TwipsPerPixelX
    UserControl.Height = 32 * Screen.TwipsPerPixelY
End Sub
