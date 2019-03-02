VERSION 5.00
Begin VB.UserControl ResizeFrame 
   Appearance      =   0  'Flat
   BackColor       =   &H00DEE2DE&
   BackStyle       =   0  'Í¸Ã÷
   ClientHeight    =   1056
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1212
   ScaleHeight     =   1056
   ScaleWidth      =   1212
   Begin VB.Label rp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   0
      Left            =   216
      TabIndex        =   0
      Top             =   192
      Width           =   120
   End
   Begin VB.Label rp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   8
      Left            =   840
      TabIndex        =   8
      Top             =   672
      Width           =   120
   End
   Begin VB.Label rp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   7
      Left            =   528
      TabIndex        =   7
      Top             =   672
      Width           =   120
   End
   Begin VB.Label rp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   6
      Left            =   216
      TabIndex        =   6
      Top             =   672
      Width           =   120
   End
   Begin VB.Label rp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   432
      Width           =   120
   End
   Begin VB.Label rp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   3
      Left            =   216
      TabIndex        =   3
      Top             =   432
      Width           =   120
   End
   Begin VB.Label rp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   192
      Width           =   120
   End
   Begin VB.Label rp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   1
      Left            =   528
      TabIndex        =   1
      Top             =   192
      Width           =   120
   End
   Begin VB.Label rp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   4
      Left            =   528
      TabIndex        =   4
      Top             =   432
      Width           =   120
   End
End
Attribute VB_Name = "ResizeFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event SizeChange()
Event Done()
Public Kid As Object, Dad As Object
Dim sx As Long, sy As Long, orr As RECT

Private Sub rp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call rp_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub rp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    Dim wr As RECT, p As POINT, cr As RECT
    GetWindowRect Dad.hwnd, wr
    GetCursorPos p
    
    p.X = (p.X - wr.Left): p.Y = (p.Y - wr.top)

    Dim r As Long
    r = rp(0).Width / Screen.TwipsPerPixelX

    cr.Left = UserControl.Extender.Left
    cr.top = UserControl.Extender.top
    cr.Right = UserControl.Extender.Width
    cr.Bottom = UserControl.Extender.Height
    
    If sx = -1 Then sx = X / Screen.TwipsPerPixelX: sy = Y / Screen.TwipsPerPixelY: orr = cr
    
    If Index = 4 Or Index = 0 Or Index = 3 Or Index = 6 Then cr.Left = p.X - sx - rp(Index).Left / Screen.TwipsPerPixelX
    If Index = 4 Or Index = 0 Or Index = 1 Or Index = 2 Then cr.top = p.Y - sy - rp(Index).top / Screen.TwipsPerPixelY
    If Index = 0 Or Index = 3 Or Index = 6 Then cr.Right = orr.Left + orr.Right - cr.Left
    If Index = 0 Or Index = 1 Or Index = 2 Then cr.Bottom = orr.top + orr.Bottom - cr.top
    If Index = 8 Or Index = 2 Or Index = 5 Then cr.Right = p.X - cr.Left + sx
    If Index = 8 Or Index = 6 Or Index = 7 Then cr.Bottom = p.Y - cr.top + sy
    
    UserControl.Extender.Left = cr.Left
    UserControl.Extender.top = cr.top
    UserControl.Extender.Width = cr.Right
    UserControl.Extender.Height = cr.Bottom
    
    Kid.Left = cr.Left + r
    Kid.top = cr.top + r
    Kid.Width = cr.Right - r * 2
    Kid.Height = cr.Bottom - r * 2
    
    RaiseEvent SizeChange
End Sub

Private Sub rp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call rp_MouseMove(Index, Button, Shift, X, Y)
    sx = -1
    RaiseEvent Done
End Sub

Private Sub UserControl_Initialize()
    sx = -1
End Sub

Private Sub UserControl_Resize()
    Dim ll As Long, lc As Long, lr As Long, bl As Long, bc As Long, br As Long
    ll = 0: lc = UserControl.Width / 2 - rp(0).Width / 2: lr = UserControl.Width - rp(0).Width
    bl = 0: bc = UserControl.Height / 2 - rp(0).Height / 2: br = UserControl.Height - rp(0).Height
    
    rp(0).Move ll, bl: rp(1).Move lc, bl: rp(2).Move lr, bl
    rp(3).Move ll, bc: rp(5).Move lr, bc
    rp(6).Move ll, br: rp(7).Move lc, br: rp(8).Move lr, br
    
    rp(4).Move lc, bc
End Sub

Public Sub RefreshPoints()
    Dim r As Long
    r = rp(0).Width / Screen.TwipsPerPixelX
    UserControl.Extender.Left = Kid.Left - r
    UserControl.Extender.top = Kid.top - r
    UserControl.Extender.Width = Kid.Width + 2 * r
    UserControl.Extender.Height = Kid.Height + 2 * r
End Sub
