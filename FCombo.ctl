VERSION 5.00
Begin VB.UserControl FCombo 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1992
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6072
   ScaleHeight     =   166
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   506
   Begin VB.Timer WOpen 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   4590
      Top             =   1260
   End
   Begin VB.ComboBox mCombo 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   10.8
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   420
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4485
   End
End
Attribute VB_Name = "FCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Made by ·½³Ì
Private mFont As Font, mCaption As String, mFontColor As OLE_COLOR, mForeColor As OLE_COLOR, mBackColor As OLE_COLOR
Private mFocus As Boolean, mPress As Boolean, mEnable As Boolean, mMouseIn As Boolean
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_CLOSE = &H10
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Event Click()

Public Sub AddItem(Item As String, Optional Index)
    mCombo.AddItem Item, Index
End Sub
Public Sub RemoveItem(Index As Integer)
    mCombo.RemoveItem Index
End Sub
Public Function ListCount() As Integer
    ListCount = mCombo.ListCount
End Function
Public Sub Clear()
    mCombo.Clear
End Sub

Private Sub mCombo_Change()
    RaiseEvent Click
End Sub

Private Sub mCombo_Click()
    mCaption = mCombo.Text
    Call UserControl_Paint
    RaiseEvent Click
End Sub
Public Property Get List(i As Integer) As String
    List = mCombo.List(i)
End Property
Public Property Let List(i As Integer, s As String)
    mCombo.List(i) = s
End Property
Public Property Get ListIndex() As Integer
    ListIndex = mCombo.ListIndex
End Property
Public Property Let ListIndex(i As Integer)
    mCombo.ListIndex = i
End Property
Property Get Text() As String
    Text = mCaption
End Property
Property Let Text(V As String)
    mCaption = V
    PropertyChanged "Caption"
    Call UserControl_Paint
End Property
Public Property Get Font() As Font
    Set Font = mFont
End Property
Public Property Set Font(f As Font)
    Set mFont = f
    Set mCombo.Font = f
    PropertyChanged "Font"
    Call UserControl_Paint
End Property
Public Property Get Fontcolor() As OLE_COLOR
    Fontcolor = mFontColor
End Property
Public Property Let Fontcolor(newColor As OLE_COLOR)
    mFontColor = newColor
    PropertyChanged "Fontcolor"
    Call UserControl_Paint
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
    mFontColor = &H606060
    mCaption = Extender.name
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mPress = True
    UserControl_Paint
    WOpen.Enabled = True
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
Private Sub UserControl_Initialize()
    SendMessage FindWindowEx(mCombo.hwnd, 0, "Edit", vbNullString), WM_CLOSE, 0, 0
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
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
    End With
End Sub
Private Sub UserControl_Paint()
    Dim Graphics As Long, ButtonDeeph As Byte
    Dim Fontfam As Long, Strformat As Long, MyFont As Long, Rclayout As RECTF, RECT As RECTF
    Dim mToken As Long, Inputbuf As GdiplusStartupInput
    Inputbuf.GdiplusVersion = 1: GdiplusStartup mToken, Inputbuf
    
    GdipCreateFromHDC UserControl.hdc, Graphics '»­²¼
    GdipSetSmoothingMode Graphics, SmoothingModeAntiAlias

    With Rclayout
        .Left = 0
        .Right = UserControl.Width \ Screen.TwipsPerPixelX
        .top = 0
        .Bottom = UserControl.Height \ Screen.TwipsPerPixelY
    End With
    GdipGraphicsClear Graphics, OLEColorChange(mBackColor)
    
    Dim DashPen As Long, LineBrush As Long
    GdipCreatePen1 OLEColorChange(mForeColor), 1, UnitPixel, DashPen
    GdipCreateLineBrush NewPointF(Rclayout.Right - 40, Rclayout.Bottom \ 2), NewPointF(Rclayout.Right - 20, Rclayout.Bottom \ 2), OLEColorChange(mBackColor, 0), OLEColorChange(mBackColor), WrapModeTile, LineBrush
       

    If Not mFont Is Nothing Then
        GdipCreateFontFamilyFromName StrPtr(mFont.name), 0, Fontfam
        GdipCreateFont Fontfam, mFont.Size, FontStyleRegular, UnitPoint, MyFont
        GdipCreateStringFormat 0, 0, Strformat
        GdipSetStringFormatAlign Strformat, StringAlignmentNear
        GdipMeasureString Graphics, StrPtr(mCaption), Len(mCaption), MyFont, Rclayout, Strformat, RECT, 0, 0
        RECT.top = (Rclayout.Bottom - RECT.Bottom) / 2
        RECT.Right = UserControl.Width
        GdipDrawString Graphics, StrPtr(mCaption), -1, MyFont, RECT, Strformat, NewBrush(OLEColorChange(mFontColor))
    End If
      
    GdipFillRectangle Graphics, LineBrush, Rclayout.Right - 40, 0, 20, Rclayout.Bottom
    GdipFillRectangle Graphics, NewBrush(OLEColorChange(mBackColor)), Rclayout.Right - 20, 0, 20, Rclayout.Bottom
    'GdipDrawRectangleI Graphics, DashPen, Rclayout.Left, Rclayout.top, Rclayout.Right - 1, Rclayout.Bottom - 1
    GdipDrawLine Graphics, DashPen, Rclayout.Right - 20, Rclayout.Bottom / 2 - 3, Rclayout.Right - 14, Rclayout.Bottom / 2 + 3
    GdipDrawLine Graphics, DashPen, Rclayout.Right - 14, Rclayout.Bottom / 2 + 3, Rclayout.Right - 8, Rclayout.Bottom / 2 - 3

    ButtonDeeph = 60
    If mMouseIn Then GdipFillRectangleI Graphics, NewBrush(OLEColorChange(mForeColor, ButtonDeeph)), 0, 0, Rclayout.Right - 1, Rclayout.Bottom - 1
  
    UserControl.Refresh
    
    GdipDeletePen DashPen
    GdipDeleteBrush LineBrush
    GdipDeleteGraphics Graphics
    GdipDeleteFontFamily Fontfam
    GdipDeleteStringFormat Strformat
    GdipDeleteFont MyFont
    
    GdiplusShutdown mToken
End Sub
Private Sub UserControl_Resize()
    mCombo.Width = UserControl.Width \ Screen.TwipsPerPixelX
    mCombo.top = UserControl.Height \ Screen.TwipsPerPixelX - mCombo.Height + 2
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
Private Sub WOpen_Timer()
    SendMessage mCombo.hwnd, WM_LBUTTONDOWN, 0, 0
    SendMessage mCombo.hwnd, WM_LBUTTONUP, 0, 0
    mMouseIn = False
    mPress = False
    UserControl_Paint
    WOpen = False
End Sub
