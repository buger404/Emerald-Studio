VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "image.ocx"
Begin VB.Form AboutWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Me"
   ClientHeight    =   5736
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   8112
   Icon            =   "AboutWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   478
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
   StartUpPosition =   1  '所有者中心
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENTION"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6144
      TabIndex        =   14
      Top             =   4692
      Width           =   1092
   End
   Begin VB.Label Label15 
      BackColor       =   &H00CEDB1A&
      Height          =   1188
      Left            =   5808
      TabIndex        =   15
      Tag             =   "switch.on"
      Top             =   4224
      Width           =   1764
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00CEDA1A&
      BorderWidth     =   3
      Tag             =   "text.title"
      X1              =   294
      X2              =   352
      Y1              =   384
      Y2              =   384
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "扁平ComboBox + Toggle + IconButton - made by 方程"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808580&
      Height          =   252
      Left            =   2304
      TabIndex        =   11
      Tag             =   "text.content"
      Top             =   3480
      Width           =   5004
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Redstone Supremacy / Inter.Net"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDA1A&
      Height          =   252
      Left            =   2304
      TabIndex        =   10
      Tag             =   "text.title"
      Top             =   3168
      Width           =   2940
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "支持:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808580&
      Height          =   252
      Left            =   504
      TabIndex        =   9
      Tag             =   "text.content"
      Top             =   3144
      Width           =   432
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "制作:                       Error 404 / 方程"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808580&
      Height          =   252
      Left            =   504
      TabIndex        =   8
      Tag             =   "text.content"
      Top             =   1560
      Width           =   3228
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Github Pages:         https://red-error404.github.io/233"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808580&
      Height          =   252
      Left            =   504
      TabIndex        =   7
      Tag             =   "text.content"
      Top             =   2736
      Width           =   4920
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "此开源项目使用MIT开源协议（包括Emerald绘图框架）"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808580&
      Height          =   252
      Left            =   648
      TabIndex        =   6
      Tag             =   "text.content"
      Top             =   5040
      Width           =   4752
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "最终代码调试和编译由Microsoft Visual Basic 6.0完成"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808580&
      Height          =   252
      Left            =   648
      TabIndex        =   5
      Tag             =   "text.content"
      Top             =   4680
      Width           =   4620
   End
   Begin ImageX.aicAlphaImage LOGO 
      Height          =   804
      Left            =   504
      Top             =   432
      Width           =   804
      _ExtentX        =   1418
      _ExtentY        =   1418
      Image           =   "AboutWindow.frx":000C
      Props           =   5
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "这只是一个代码生成器和设计器，不是IDE"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDA1A&
      Height          =   252
      Left            =   648
      TabIndex        =   4
      Tag             =   "text.title"
      Top             =   4320
      Width           =   3576
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "邮箱:                       ris_vb@126.com"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808580&
      Height          =   288
      Left            =   504
      TabIndex        =   3
      Tag             =   "text.content"
      Top             =   2328
      Width           =   3348
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "QQ:                        1361778219 (Error 404) / 937697555 (方程)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808580&
      Height          =   252
      Left            =   504
      TabIndex        =   2
      Tag             =   "text.content"
      Top             =   1944
      Width           =   5628
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00CEDB1A&
      Caption         =   "Indev.309"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6456
      TabIndex        =   1
      Tag             =   "switch.on"
      Top             =   336
      Width           =   1224
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Emerald Studio"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDA1A&
      Height          =   252
      Left            =   1704
      TabIndex        =   0
      Tag             =   "text.title"
      Top             =   336
      Width           =   1392
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Height          =   1188
      Left            =   480
      TabIndex        =   13
      Tag             =   "window.background"
      Top             =   4224
      Width           =   7068
   End
   Begin VB.Label background 
      BackColor       =   &H00F0F2F0&
      Height          =   4860
      Left            =   0
      TabIndex        =   12
      Tag             =   "window.tool"
      Top             =   888
      Width           =   8124
   End
End
Attribute VB_Name = "AboutWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    UpdateSkin Me, CurrentSkin
End Sub

