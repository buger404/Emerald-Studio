VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "image.ocx"
Begin VB.Form AboutWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Me"
   ClientHeight    =   7224
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   7920
   Icon            =   "AboutWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   602
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   1  '所有者中心
   Begin VB.Line Line2 
      BorderColor     =   &H00CEDA1A&
      BorderWidth     =   3
      Tag             =   "text.title"
      X1              =   292
      X2              =   350
      Y1              =   474
      Y2              =   474
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
      Left            =   2280
      TabIndex        =   11
      Tag             =   "text.content"
      Top             =   4392
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
      Left            =   2280
      TabIndex        =   10
      Tag             =   "text.title"
      Top             =   4080
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
      Left            =   480
      TabIndex        =   9
      Tag             =   "text.content"
      Top             =   4080
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
      Left            =   480
      TabIndex        =   8
      Tag             =   "text.content"
      Top             =   2640
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
      Left            =   480
      TabIndex        =   7
      Tag             =   "text.content"
      Top             =   3720
      Width           =   4920
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "* 此开源项目使用MIT开源协议（包括Emerald绘图框架）"
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
      Left            =   480
      TabIndex        =   6
      Tag             =   "text.content"
      Top             =   6120
      Width           =   4956
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "* 最终代码调试和编译由Microsoft Visual Basic 6.0完成"
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
      Left            =   480
      TabIndex        =   5
      Tag             =   "text.content"
      Top             =   5760
      Width           =   4860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Tag             =   "line"
      X1              =   20
      X2              =   640
      Y1              =   440
      Y2              =   440
   End
   Begin ImageX.aicAlphaImage LOGO 
      Height          =   960
      Left            =   3480
      Top             =   600
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Image           =   "AboutWindow.frx":000C
      Props           =   5
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "* 这只是一个代码生成器和设计器，不是IDE"
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
      Height          =   288
      Left            =   480
      TabIndex        =   4
      Tag             =   "text.title"
      Top             =   5400
      Width           =   3780
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
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Tag             =   "text.content"
      Top             =   3360
      Width           =   3345
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
      Left            =   480
      TabIndex        =   2
      Tag             =   "text.content"
      Top             =   3000
      Width           =   5628
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00808580&
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Tag             =   "text.content"
      Top             =   2040
      Width           =   7875
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Tag             =   "text.title"
      Top             =   1680
      Width           =   7875
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

