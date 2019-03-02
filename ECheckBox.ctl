VERSION 5.00
Begin VB.UserControl ECheckBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
End
Attribute VB_Name = "ECheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mBackColor As Long, mOffColor As Long, mOnColor As Long, mForeColor As Long, mIsOn As Boolean
Dim mFont As StdFont, mContent As String
Public Property Get IsOn() As Boolean
    IsOn = mIsOn
End Property
Public Property Let IsOn(o As Boolean)
    mIsOn = o
    Call UserControl_Paint
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property
Public Property Let ForeColor(c As OLE_COLOR)
    mForeColor = c
    Call UserControl_Paint
End Property
Public Property Get OffColor() As OLE_COLOR
    OffColor = mOffColor
End Property
Public Property Let OffColor(c As OLE_COLOR)
    mOffColor = c
    Call UserControl_Paint
End Property
Public Property Get OnColor() As OLE_COLOR
    OnColor = mOnColor
End Property
Public Property Let OnColor(c As OLE_COLOR)
    mOnColor = c
    Call UserControl_Paint
End Property
Public Property Get Backcolor() As OLE_COLOR
    Backcolor = mBackColor
End Property
Public Property Let Backcolor(c As OLE_COLOR)
    mBackColor = c
    Call UserControl_Paint
End Property
Public Property Get Font() As StdFont
    Set Font = mFont
End Property
Public Property Set Font(f As StdFont)
    Set mFont = f
    Call UserControl_Paint
End Property
Public Property Get Content() As String
    Content = mContent
End Property
Public Property Let Content(c As String)
    mContent = c
    Call UserControl_Paint
End Property
Private Function RGBtoARGB(rgbL As Long, alpha As Long) As Long
    Dim buffer(3) As Byte, ret As Long
    CopyMemory buffer(1), rgbL, 3: buffer(0) = alpha
    CopyMemory ByVal VarPtr(ret), buffer(3), 1
    CopyMemory ByVal VarPtr(ret) + 1, buffer(2), 1
    CopyMemory ByVal VarPtr(ret) + 2, buffer(1), 1
    CopyMemory ByVal VarPtr(ret) + 3, buffer(0), 1
    
    RGBtoARGB = ret
End Function
Private Sub UserControl_Click()
    mIsOn = Not mIsOn
    Call UserControl_Paint
End Sub
Private Sub FillRoundRect(X As Long, Y As Long, w As Long, h As Long, r As Long, b As Long, g As Long)
    Dim p As Long
    GdipCreatePath FillModeWinding, p
    GdipAddPathArc p, X, Y, r, r, 180, 90
    GdipAddPathArc p, X + w - r - 1, Y, r, r, 270, 90
    GdipAddPathArc p, X + w - r - 1, Y + h - r - 1, r, r, 0, 90
    GdipAddPathArc p, X, Y + h - r - 1, r, r, 90, 90
    GdipClosePathFigure p
    GdipFillPath g, b, p
    GdipDeletePath p
End Sub
Private Sub UserControl_InitProperties()
    Dim dFont As New StdFont
    With dFont
        .name = "Î¢ÈíÑÅºÚ"
        .Size = 10
    End With
    Set mFont = dFont
    
    mBackColor = RGB(255, 255, 255): mOffColor = RGB(212, 212, 212): mOnColor = RGB(26, 218, 207)
    mForeColor = RGB(192, 192, 192)
    
    mContent = Extender.name
End Sub
Private Sub UserControl_Paint()
    Dim b As Long, g As Long, f As Long, ff As Long, strF As Long
    Dim w As Long, h As Long, r As RECTF
    Dim mToken As Long, Inputbuf As GdiplusStartupInput
    Inputbuf.GdiplusVersion = 1: GdiplusStartup mToken, Inputbuf
    
    GdipCreateFromHDC UserControl.hdc, g
    GdipSetSmoothingMode g, SmoothingModeAntiAlias
    
    w = UserControl.Width / Screen.TwipsPerPixelX
    h = UserControl.Height / Screen.TwipsPerPixelY
    
    GdipCreateSolidFill 0, b
    GdipCreateFontFamilyFromName StrPtr(mFont.name), 0, ff
    GdipCreateFont ff, mFont.Size, FontStyleRegular, UnitPoint, f
    GdipCreateStringFormat 0, 0, strF
    GdipSetStringFormatAlign strF, StringAlignmentNear
    
    GdipGraphicsClear g, RGBtoARGB(mBackColor, 255)
    
    If mIsOn Then
        GdipSetSolidFillColor b, RGBtoARGB(OnColor, 120)
        FillRoundRect 0, h / 2 - h / 3, h * 2.3 - 1, h / 3 * 2, h / 3 * 2, b, g
        GdipSetSolidFillColor b, RGBtoARGB(OnColor, 255)
        GdipFillEllipse g, b, h * 2.3 - h, 0, h - 1, h - 1
    Else
        GdipSetSolidFillColor b, RGBtoARGB(OffColor, 120)
        FillRoundRect 0, h / 2 - h / 3, h * 2.3 - 1, h / 3 * 2, h / 3 * 2, b, g
        GdipSetSolidFillColor b, RGBtoARGB(OffColor, 255)
        GdipFillEllipse g, b, 0, 0, h - 1, h - 1
    End If
    
    GdipMeasureString g, StrPtr(mContent), Len(mContent), f, NewRectF(0, 0, w, h), strF, r, 0, 0
    With r
        .Left = h * 2.3 + 10
        .top = Int(h / 2 - .Bottom / 2) - 2
    End With
    
    GdipSetSolidFillColor b, RGBtoARGB(mForeColor, 255)
    GdipDrawString g, StrPtr(mContent), -1, f, r, strF, b
    
    GdipDeleteBrush b
    GdipDeleteFont f
    GdipDeleteFontFamily ff
    GdipDeleteStringFormat strF
    GdipDeleteGraphics g
    
    GdiplusShutdown mToken
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mBackColor = Val(PropBag.ReadProperty("BackColor", RGB(255, 255, 255)))
    mOffColor = Val(PropBag.ReadProperty("OffColor", RGB(212, 212, 212)))
    mOnColor = Val(PropBag.ReadProperty("OnColor", RGB(26, 218, 207)))
    mContent = PropBag.ReadProperty("Content", Extender.name)
    mIsOn = PropBag.ReadProperty("IsOn", False)
    mForeColor = PropBag.ReadProperty("ForeColor", RGB(192, 192, 192))
    
    Dim dFont As New StdFont
    With dFont
        .name = "Î¢ÈíÑÅºÚ"
        .Size = 10
    End With
    
    Set mFont = PropBag.ReadProperty("Font", dFont)
    
    Call UserControl_Paint
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", mBackColor, RGB(255, 255, 255)
    PropBag.WriteProperty "OffColor", mOffColor, RGB(212, 212, 212)
    PropBag.WriteProperty "OnColor", mOnColor, RGB(26, 218, 207)
    PropBag.WriteProperty "Content", mContent, Extender.name
    PropBag.WriteProperty "IsOn", mIsOn
    PropBag.WriteProperty "ForeColor", mForeColor
    
    Dim dFont As New StdFont
    With dFont
        .name = "Î¢ÈíÑÅºÚ"
        .Size = 10
    End With
    
    PropBag.WriteProperty "Font", mFont, dFont
End Sub
