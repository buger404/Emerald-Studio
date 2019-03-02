VERSION 5.00
Begin VB.Form StartupWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "StartupWindow"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9996
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   833
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Waiter 
      Interval        =   1
      Left            =   360
      Top             =   240
   End
End
Attribute VB_Name = "StartupWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dc As Long, b As BLENDFUNCTION, ws As size, p As POINTAPI

'设置窗口图像
'<path:图像路径>
Public Sub UploadLOGO(ByVal path As String)
    Dim i As Long, w As Long, h As Long, g As Long
    GdipCreateBitmapFromFile StrPtr(path), i
    GdipGetImageWidth i, w: GdipGetImageHeight i, h
    w = w * 0.8: h = h * 0.8
    
    dc = CreateCDC(w, h)
    GdipCreateFromHDC dc, g
    GdipDrawImageRect g, i, 0, 0, w, h
    
    GdipDisposeImage i
    GdipDeleteGraphics g
    
    SetWindowLongA Me.hwnd, GWL_EXSTYLE, GetWindowLongA(Me.hwnd, -20) Or &H80000
    
    ws.cx = w: ws.cy = h: p.X = 0: p.Y = 0
    With b
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    UpdateLayeredWindow Me.hwnd, 0, 0, ws, dc, 0, 0, b, &H2
    
    Me.Move Screen.Width / 2 - (w * Screen.TwipsPerPixelX) / 2, Screen.Height / 2 - (h * Screen.TwipsPerPixelY) / 2, w * Screen.TwipsPerPixelX, h * Screen.TwipsPerPixelY
End Sub

Private Sub Form_Load()
    Me.Show
    UploadLOGO App.path & "\assets\logo.png"
End Sub

Private Sub Waiter_Timer()
    If Data.GetData("IndevWarn") = "" Then '如果还没警告用户，则警告
        WarnWindow.Show
    Else
        ProjectWindow.Show
    End If
    Unload Me
End Sub
