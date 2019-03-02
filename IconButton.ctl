VERSION 5.00
Begin VB.UserControl IconButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "IconButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'图标按钮
'Made by 方程
'—————————————————————————————————————————————————————————————————————————————————
'属性：  IconPath 图标的绝对路径
'                 例如使用Assets文件夹下的img.png：Assets\img.png
'       IconRatio 要显示的图标占整个按钮的大小的比例，会根据图标长宽尺寸缩放
'—————————————————————————————————————————————————————————————————————————————————

Private mFont As Font, mForeColor As OLE_COLOR, mBackColor As OLE_COLOR, mIconPath As String, mIconRatio As Single
Private mFocus As Boolean, mPress As Boolean, mEnable As Boolean, mMouseIn As Boolean
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Event Click()
Public Property Get IconRatio() As Single
    IconRatio = mIconRatio
End Property
Public Property Let IconRatio(Value As Single)
    mIconRatio = Value
    PropertyChanged "IconRatio"
    Call UserControl_Paint
End Property
Public Property Let IconPath(RelativePath As String)
    mIconPath = RelativePath
    PropertyChanged "Iconpath"
    Call UserControl_Paint
End Property
Public Property Get IconPath() As String
    IconPath = mIconPath
End Property
Public Property Get Backcolor() As OLE_COLOR
    Backcolor = mBackColor
End Property
Public Property Let Backcolor(newColor As OLE_COLOR)
    mBackColor = newColor
    PropertyChanged "Backcolor"
    Call UserControl_Paint
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property
Public Property Let ForeColor(newColor As OLE_COLOR)
    mForeColor = newColor
    PropertyChanged "Forecolor"
    Call UserControl_Paint
End Property
Private Sub UserControl_InitProperties()
    Set mFont = Ambient.Font
    mFont.Size = 16
    mBackColor = RGB(225, 225, 225)
    mForeColor = &H808080
    mCaption = Extender.name
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mPress = True
    UserControl_Paint
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMouseIn = False Then
        SetCapture UserControl.hwnd
        mMouseIn = True
        Call UserControl_Paint
    ElseIf mMouseIn And (X > UserControl.Width \ Screen.TwipsPerPixelX Or Y > UserControl.Height \ Screen.TwipsPerPixelY Or X < 0 Or Y < 0) Then
        ReleaseCapture
        mMouseIn = False
        Call UserControl_Paint
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    mPress = False
    mMouseIn = False
    Call UserControl_Paint
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        Set mFont = .ReadProperty("Font")
        Set mCombo.Font = mFont
        mBackColor = .ReadProperty("Backcolor")
        mFontColor = .ReadProperty("Fontcolor")
        mForeColor = .ReadProperty("Forecolor")
        mCaption = .ReadProperty("Caption")
        mIconPath = .ReadProperty("IconPath")
        mIconRatio = .ReadProperty("IconRatio")
    End With
    Call UserControl_Paint
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Font", mFont, Ambient.Font
        .WriteProperty "Backcolor", mBackColor, 0
        .WriteProperty "Forecolor", mForeColor, 0
        .WriteProperty "Fontcolor", mFontColor, 0
        .WriteProperty "Caption", mCaption, 0
        .WriteProperty "IconPath", mIconPath, 0
        .WriteProperty "IconRatio", mIconRatio, 0
    End With
End Sub
Private Sub UserControl_Paint()
    Dim Graphics As Long, ButtonDeeph As Byte
    Dim Img As Long
    Dim Rclayout As RECTF, ImgRect As RECTF, ImgW As Long, ImgH As Long
    Dim mToken As Long, Inputbuf As GdiplusStartupInput
    Inputbuf.GdiplusVersion = 1: GdiplusStartup mToken, Inputbuf
    
    GdipCreateFromHDC UserControl.hdc, Graphics '画布
    GdipSetSmoothingMode Graphics, SmoothingModeAntiAlias
    
    With Rclayout
        .Left = 0
        .Right = UserControl.Width \ Screen.TwipsPerPixelX + 1
        .top = 0
        .Bottom = UserControl.Height \ Screen.TwipsPerPixelY + 1
    End With
    
    GdipGraphicsClear Graphics, OLEColorChange(mBackColor)
    
    Dim DashPen As Long, LineBrush As Long
    GdipCreatePen1 OLEColorChange(mForeColor), 1, UnitPixel, DashPen
    GdipCreateLineBrush NewPointF(Rclayout.Right - 40, Rclayout.Bottom \ 2), NewPointF(Rclayout.Right - 20, Rclayout.Bottom \ 2), OLEColorChange(mBackColor, 0), OLEColorChange(mBackColor), WrapModeTile, LineBrush
    GdipLoadImageFromFile StrPtr(App.path & "\" & mIconPath), Img
    If Img <> 0 Then
        GdipGetImageWidth Img, ImgW
        GdipGetImageHeight Img, ImgH
        
        With ImgRect
            .Right = Rclayout.Right * mIconRatio
            .Bottom = ImgH / ImgW * ImgRect.Right
            .Left = (Rclayout.Right - ImgRect.Right) / 2
            .top = (Rclayout.Bottom - ImgRect.Right) / 2
        End With
    End If
    
    GdipFillRectangle Graphics, LineBrush, Rclayout.Right - 40, 0, 20, Rclayout.Bottom
    
    '************************************************************
    '这里NewBrush会造成内存泄漏 特此标记 来日修复
    'FCombo也需要修复
    '************************************************************
    
    GdipFillRectangle Graphics, NewBrush(OLEColorChange(mBackColor)), Rclayout.Right - 20, 0, 20, Rclayout.Bottom
    GdipDrawImageRect Graphics, Img, ImgRect.Left, ImgRect.top, ImgRect.Right, ImgRect.Bottom

    ButtonDeeph = 60
    If mMouseIn Then GdipFillRectangleI Graphics, NewBrush(OLEColorChange(mForeColor, ButtonDeeph)), -1, -1, Rclayout.Right, Rclayout.Bottom
    If mPress Then GdipFillRectangleI Graphics, NewBrush(OLEColorChange(mForeColor, ButtonDeeph)), -1, -1, Rclayout.Right - 1, Rclayout.Bottom - 1
    UserControl.Refresh
    
    GdipDeletePen DashPen
    GdipDeleteBrush LineBrush
    GdipDeleteGraphics Graphics
    GdipDisposeImage Img
    GdiplusShutdown mToken
End Sub
Private Sub UserControl_Resize()
    Call UserControl_Paint
End Sub
Public Function OLEColorChange(Color As OLE_COLOR, Optional ColorAlpha As Byte = 255) As Long
    Dim c, i, Ccount
    c = Hex(Color)
    If Len(c) < 6 Then
        Ccount = 6 - Len(c)
        For i = 1 To Ccount
         c = "0" & c
        Next
    End If
    OLEColorChange = "&h" & Hex(ColorAlpha) & Mid(c, 5, 2) & Mid(c, 3, 2) & Mid(c, 1, 2)
End Function


