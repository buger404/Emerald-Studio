VERSION 5.00
Begin VB.UserControl EDsnObject 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  '透明
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ClipBehavior    =   0  '无
   ScaleHeight     =   3540
   ScaleWidth      =   4500
   Windowless      =   -1  'True
   Begin VB.Label tiptext 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "请先初始化你的GDIPlus再继续"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1176
      TabIndex        =   0
      Top             =   1344
      Visible         =   0   'False
      Width           =   2460
   End
End
Attribute VB_Name = "EDsnObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public lWin As Long, lObj As Long, DrawLock As Boolean
Public Sub Refresh()
    UserControl.Refresh
End Sub
Private Sub UserControl_Paint()
    '参考自方程的百度帖子
    'url=http://tieba.baidu.com/p/4719499280
    
    Dim wg As Long, g As Long, mi As Long, w As Long, h As Long
    If Not GdipIsStarted Then
        tiptext.Visible = True
        tiptext.Move 10, 10
        Exit Sub
    End If
    If lWin = 0 Then Exit Sub
    
    w = UserControl.Width / Screen.TwipsPerPixelX: h = UserControl.Height / Screen.TwipsPerPixelY
    
    GdipCreateFromHDC UserControl.hdc, wg
    CreateBitmapWithGraphics mi, g, w, h, PixelFormat32bppARGB
    GdipGraphicsClear g, 0
    GdipSetSmoothingMode g, SmoothingModeAntiAlias
    GdipSetTextRenderingHint g, TextRenderingHintAntiAlias
    
    If DrawLock Then
        Dim p2 As Long, b2 As Long
        GdipSetSmoothingMode g, SmoothingModeDefault
        GdipCreatePen1 argb(255, 26, 219, 206), 1, UnitPixel, p2
        GdipCreateSolidFill argb(120, 26, 219, 206), b2
        GdipFillRectangle g, b2, 0, 0, w, h
        GdipDrawRectangle g, p2, 0, 0, w - 1, h - 1
        GdipDeleteBrush b2: GdipDeletePen p2
        GoTo last
    End If
    
    With DsnWin(lWin).Obj(lObj)
        Select Case .kind
            Case 0
                Dim i As Long
                GdipCreateBitmapFromFile StrPtr(App.path & "\assets\bm_ok.png"), i
                GdipDrawImageRect g, i, 0, 0, w, h
                GdipDisposeImage i
            Case 1
                EF.Writes .Content, 0, 0, g, .Color, .size, w, h, .align, .style
            Case 2
                Dim b As Long
                GdipCreateSolidFill .Color, b
                If .Content = 0 Then
                    GdipFillRectangle g, b, 0, 0, w, h
                Else
                    GdipFillEllipse g, b, 0, 0, w, h
                End If
                GdipDeleteBrush b
        End Select
    End With
    
last:
    GdipDrawImage wg, mi, 0, 0
    GdipDeleteGraphics g
    GdipDisposeImage mi
    GdipDeleteGraphics wg
    
End Sub
