VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "image.ocx"
Begin VB.Form SkinTest 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   0  'None
   Caption         =   "SkinTest"
   ClientHeight    =   6156
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   7368
   ForeColor       =   &H008C8080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   614
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin Emerald_Studio.ResizeFrame ResizeFrame1 
      Height          =   756
      Left            =   3144
      TabIndex        =   22
      Top             =   648
      Width           =   2508
      _ExtentX        =   6244
      _ExtentY        =   1545
   End
   Begin Emerald_Studio.IconButton IconButton1 
      Height          =   492
      Left            =   1656
      TabIndex        =   21
      Top             =   1584
      Width           =   828
      _ExtentX        =   1461
      _ExtentY        =   868
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   14803425
      Forecolor       =   8421504
      IconPath        =   "assets\Img.png"
      IconRatio       =   0.5
   End
   Begin Emerald_Studio.Toggle Toggle1 
      Height          =   468
      Left            =   504
      TabIndex        =   20
      Top             =   1584
      Width           =   732
      _ExtentX        =   1291
      _ExtentY        =   826
      Backcolor       =   14803425
      Forecolor       =   8421504
      IconPath        =   "assets\Img.png"
      IconRatio       =   0.5
   End
   Begin Emerald_Studio.EButton EButton1 
      Height          =   372
      Left            =   1320
      TabIndex        =   5
      Top             =   5112
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   656
      HoverColor      =   15592941
      ForeColor       =   10000536
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "ÆÕÍ¨°´Å¥"
      Align           =   0
   End
   Begin Emerald_Studio.EButton toolbutton 
      Height          =   372
      Index           =   0
      Left            =   2976
      TabIndex        =   6
      Top             =   2640
      Width           =   8052
      _ExtentX        =   14203
      _ExtentY        =   656
      DefaultColor    =   15592941
      HoverColor      =   13619151
      ForeColor       =   10000536
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "¼òµ¥°´Å¥"
      Align           =   1
   End
   Begin Emerald_Studio.EButton okbtn 
      Height          =   372
      Left            =   2736
      TabIndex        =   7
      Top             =   5112
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   656
      HoverColor      =   15727333
      ForeColor       =   7457838
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "¸ßÁÁ°´Å¥"
      Align           =   0
   End
   Begin Emerald_Studio.EButton toolitems 
      Height          =   492
      Index           =   0
      Left            =   2976
      TabIndex        =   9
      Top             =   3144
      Width           =   3132
      _ExtentX        =   5525
      _ExtentY        =   868
      HoverColor      =   15592941
      ForeColor       =   10000536
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "¹¤¾ß°´Å¥"
      Align           =   1
   End
   Begin Emerald_Studio.ECheckBox forcheck 
      Height          =   228
      Index           =   0
      Left            =   2976
      TabIndex        =   15
      Top             =   1248
      Width           =   1428
      _ExtentX        =   2519
      _ExtentY        =   402
      BackColor       =   15592941
      OffColor        =   9408399
      OnColor         =   13556506
      Content         =   "¹Ø"
      IsOn            =   0   'False
      ForeColor       =   10000536
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Emerald_Studio.ECheckBox forcheck 
      Height          =   228
      Index           =   1
      Left            =   4608
      TabIndex        =   16
      Top             =   1248
      Width           =   1428
      _ExtentX        =   2519
      _ExtentY        =   402
      BackColor       =   15592941
      OffColor        =   12632256
      OnColor         =   7457838
      Content         =   "¿ª"
      IsOn            =   -1  'True
      ForeColor       =   10000536
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Emerald_Studio.FCombo objCombo 
      Height          =   324
      Left            =   2928
      TabIndex        =   17
      Top             =   840
      Width           =   3156
      _ExtentX        =   5567
      _ExtentY        =   572
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   16777215
      Forecolor       =   10000536
      Fontcolor       =   10000536
      Caption         =   "¹¤¾ßÏÂÀ­¿ò"
   End
   Begin Emerald_Studio.FCombo FCombo1 
      Height          =   324
      Left            =   2928
      TabIndex        =   18
      Top             =   480
      Width           =   3156
      _ExtentX        =   5567
      _ExtentY        =   572
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   15592941
      Forecolor       =   10000536
      Fontcolor       =   10000536
      Caption         =   "ÆÕÍ¨ÏÂÀ­¿ò"
   End
   Begin Emerald_Studio.EEdit EEdit1 
      Height          =   372
      Left            =   144
      TabIndex        =   19
      Top             =   4008
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   656
      BackColor       =   15592941
      ForeColor       =   10000536
      BorderColor     =   7457838
      Alignment       =   0
      LockInput       =   0   'False
   End
   Begin VB.Label tooldes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ÆÕÍ¨ÎÄ±¾"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00989898&
      Height          =   252
      Index           =   0
      Left            =   3000
      TabIndex        =   14
      Top             =   1560
      Width           =   768
   End
   Begin VB.Shape prepareFrame 
      BackColor       =   &H00E6F7D9&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0071CC2E&
      Height          =   612
      Left            =   3000
      Top             =   3720
      Width           =   1884
   End
   Begin VB.Label toolpad 
      Appearance      =   0  'Flat
      BackColor       =   &H00302D2D&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FCFDF0&
      Height          =   450
      Index           =   2
      Left            =   1560
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label toolpad 
      Appearance      =   0  'Flat
      BackColor       =   &H00422E00&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FCFDF0&
      Height          =   450
      Index           =   1
      Left            =   960
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label toolpad 
      Appearance      =   0  'Flat
      BackColor       =   &H003F4307&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FCFDF0&
      Height          =   450
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DEE2DE&
      ForeColor       =   &H00989898&
      Height          =   492
      Left            =   2160
      TabIndex        =   10
      Top             =   696
      Visible         =   0   'False
      Width           =   468
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00645959&
      Height          =   405
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7530
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B6B6B6&
      X1              =   74
      X2              =   530
      Y1              =   388
      Y2              =   388
   End
   Begin VB.Label UIOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0071CC2E&
      Caption         =   "Ñ¡ÖÐ"
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
      Height          =   312
      Index           =   0
      Left            =   4224
      TabIndex        =   4
      Top             =   5160
      Width           =   1320
   End
   Begin VB.Label UIOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008F8F8F&
      Caption         =   "Î´Ñ¡ÖÐ"
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
      Height          =   312
      Index           =   1
      Left            =   5760
      TabIndex        =   3
      Top             =   5160
      Width           =   1320
   End
   Begin VB.Label tooldes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "¸±±êÌâÎÄ±¾"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Index           =   1
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   960
   End
   Begin VB.Label title 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "±êÌâÎÄ±¾"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0071CC2E&
      Height          =   252
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Width           =   768
   End
   Begin VB.Label background 
      Appearance      =   0  'Flat
      BackColor       =   &H0071CC2E&
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
      Height          =   405
      Left            =   -120
      TabIndex        =   0
      Top             =   5760
      Width           =   7530
   End
   Begin ImageX.aicAlphaImage LOGO 
      Height          =   960
      Left            =   912
      Top             =   2352
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      Image           =   "SkinTest.frx":0000
      Props           =   5
   End
End
Attribute VB_Name = "SkinTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set ResizeFrame1.Dad = Me
    Set ResizeFrame1.Kid = forcheck(0)
End Sub
