VERSION 5.00
Begin VB.Form dsnWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6468
   ControlBox      =   0   'False
   Icon            =   "dsnWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   539
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox tempBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   708
      Left            =   4848
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   5016
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox titleframe 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   768
      ScaleWidth      =   6204
      TabIndex        =   2
      Top             =   0
      Width           =   6204
      Begin VB.Label title 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Page1"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9.6
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   570
      End
      Begin VB.Label background 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEDA1A&
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9.6
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CEDA1A&
         Height          =   765
         Left            =   0
         TabIndex        =   4
         Tag             =   "window.highlight"
         Top             =   0
         Width           =   6210
      End
   End
   Begin VB.PictureBox us 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F2F2F2&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   0
      Left            =   1080
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Emerald_Studio.ResizeFrame rframe 
      Height          =   756
      Left            =   3456
      TabIndex        =   1
      Top             =   1368
      Visible         =   0   'False
      Width           =   2508
      _ExtentX        =   6244
      _ExtentY        =   1545
   End
   Begin VB.Shape frames 
      BackColor       =   &H00F9F9F9&
      BorderColor     =   &H00CEDA1A&
      Height          =   6252
      Left            =   24
      Tag             =   "window.highlight"
      Top             =   0
      Width           =   6132
   End
   Begin VB.Shape prepareFrame 
      BackColor       =   &H00F8FAD8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CEDA1A&
      Height          =   2775
      Left            =   1080
      Tag             =   "focus"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "dsnWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PrintWindow Lib "user32" (ByVal SrcHwnd As Long, ByVal DesHDC As Long, ByVal uFlag As Long) As Long
Dim nPage As String, nindex As Integer
Dim sx As Long, sy As Long, ex As Long, ey As Long
Public Property Get PageName() As String
    PageName = nPage
End Property
Public Property Let PageName(n As String)
    nPage = n: title.Caption = n
    MainWindow.pagelist.List(nindex - 1) = n
End Property

Private Sub Form_Click()
    rframe.Visible = False
End Sub

Private Sub Form_Load()
    Set rframe.Dad = Me
    nPage = "Page1"
    UpdateSkin Me, CurrentSkin
    Me.Backcolor = RGB(255, 255, 255)
    AddDsnWindow Me: nindex = UBound(DsnWin)
    MainWindow.pagelist.AddItem nPage
    sx = -1
    ReDim objs(0)
End Sub
Public Sub UpdateUS(ByVal Index As Integer)
    Dim g As Long, w As Long, h As Long
    w = us(Index).Width: h = us(Index).Height
    
    us(Index).Visible = False
    PrintWindow Me.hwnd, tempBox.hdc, vbSrcCopy
    tempBox.Refresh
    BitBlt us(Index).hdc, 0, 0, us(Index).Width, us(Index).Height, tempBox.hdc, us(Index).Left, us(Index).top, vbSrcCopy
    us(Index).Refresh
    
    GdipCreateFromHDC us(Index).hdc, g
    GdipSetSmoothingMode g, SmoothingModeAntiAlias

    With DsnWin(nindex).Obj(Index)
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
    
    GdipDeleteGraphics g
    
    us(Index).Visible = True
    us(Index).Refresh
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Not Droping) Or (Button = 0) Then Exit Sub
    
    If sx = -1 Then
        sx = X: sy = Y: prepareFrame.Visible = True
    End If
    
    Dim tsx As Long, tsy As Long, tex As Long, tey As Long
    If X < sx Then
        tsx = X: tex = sx
    Else
        tsx = sx: tex = X
    End If
    If Y < sy Then
        tsy = Y: tey = sy
    Else
        tsy = sy: tey = Y
    End If
    ex = X: ey = Y
    
    prepareFrame.Move tsx, tsy, (tex - tsx), (tey - tsy)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Droping Then
        Dim tsx As Long, tsy As Long, tex As Long, tey As Long
        If X < sx Then
            tsx = X: tex = sx
        Else
            tsx = sx: tex = X
        End If
        If Y < sy Then
            tsy = Y: tey = sy
        Else
            tsy = sy: tey = Y
        End If
        
        Load us(us.UBound + 1)
        With us(us.UBound)
            .Move tsx, tsy, tex - tsx, tey - tsy
            .Visible = True
            .ZOrder
        End With
        
        ReDim Preserve DsnWin(nindex).Obj(UBound(DsnWin(nindex).Obj) + 1)
        With DsnWin(nindex).Obj(UBound(DsnWin(nindex).Obj))
            .kind = DropI
            .Color = argb(255, 27, 27, 27)
            .size = 16
            Select Case .kind
                Case 0
                Case 1
                    .Content = "text"
                Case 2
                    .Content = 0
            End Select
            Set .pad = us(us.UBound)
        End With
        
        If MainWindow.nPage = nindex Then
            MainWindow.objCombo.AddItem "[Object " & UBound(DsnWin(nindex).Obj) & " ]"
            MainWindow.objCombo.ListIndex = MainWindow.objCombo.ListCount - 1
        End If
        
        Screen.MousePointer = 0: Droping = False: sx = -1: prepareFrame.Visible = False
        
        Call UpdateUS(us.UBound)
        Call SetFocusIn(us.UBound)
    End If
End Sub

Public Sub SetFocusIn(i As Integer)
    Set rframe.Kid = us(i)
    rframe.ZOrder
    titleframe.ZOrder
    rframe.Visible = True
    rframe.RefreshPoints
End Sub

Private Sub Form_Resize()
    frames.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    background.Width = Me.Width
    titleframe.Width = Me.ScaleWidth
    tempBox.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub rframe_Done()
    UpdateUS rframe.Kid.Index
    '¸üÐÂÊôÐÔ±í
    MainWindow.protext(0).Content = us(rframe.Kid.Index).Left
    MainWindow.protext(1).Content = us(rframe.Kid.Index).top - background.Height / Screen.TwipsPerPixelY
    MainWindow.protext(2).Content = us(rframe.Kid.Index).Width
    MainWindow.protext(3).Content = us(rframe.Kid.Index).Height
End Sub

Private Sub us_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MainWindow.objCombo.ListIndex = Index - 1
    Call SetFocusIn(Index)
End Sub
