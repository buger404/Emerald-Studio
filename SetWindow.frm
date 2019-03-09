VERSION 5.00
Begin VB.Form SetWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emerald Studio"
   ClientHeight    =   6180
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   9876
   Icon            =   "SetWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   515
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   823
   StartUpPosition =   2  '屏幕中心
   Begin Emerald_Studio.EEdit Text4 
      Height          =   288
      Left            =   3120
      TabIndex        =   22
      Tag             =   "edit"
      Top             =   5448
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   508
      Content         =   "120"
      ForeColor       =   9871510
      BorderColor     =   13556506
      Alignment       =   0
      LockInput       =   0   'False
   End
   Begin Emerald_Studio.EEdit Text3 
      Height          =   288
      Left            =   3144
      TabIndex        =   20
      Tag             =   "edit"
      Top             =   5136
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   508
      Content         =   "10"
      ForeColor       =   9871510
      BorderColor     =   13556506
      Alignment       =   0
      LockInput       =   0   'False
   End
   Begin Emerald_Studio.EEdit Text2 
      Height          =   288
      Left            =   1728
      TabIndex        =   15
      Tag             =   "edit"
      Top             =   2928
      Width           =   6132
      _ExtentX        =   0
      _ExtentY        =   0
      Content         =   "未设置"
      ForeColor       =   9871510
      BorderColor     =   13556506
      Alignment       =   0
      LockInput       =   0   'False
   End
   Begin Emerald_Studio.EEdit Text1 
      Height          =   288
      Left            =   1728
      TabIndex        =   12
      Tag             =   "edit"
      Top             =   2568
      Width           =   6132
      _ExtentX        =   0
      _ExtentY        =   0
      Content         =   "未设置"
      ForeColor       =   9871510
      BorderColor     =   13556506
      Alignment       =   0
      LockInput       =   0   'False
   End
   Begin Emerald_Studio.EEdit pathtext 
      Height          =   288
      Left            =   1728
      TabIndex        =   4
      Tag             =   "edit"
      Top             =   1464
      Width           =   6132
      _ExtentX        =   0
      _ExtentY        =   0
      Content         =   "未设置"
      ForeColor       =   9871510
      BorderColor     =   13556506
      Alignment       =   0
      LockInput       =   0   'False
   End
   Begin Emerald_Studio.EButton pather 
      Height          =   252
      Left            =   8064
      TabIndex        =   3
      Tag             =   "button:default"
      Top             =   1512
      Width           =   372
      _ExtentX        =   677
      _ExtentY        =   445
      DefaultColor    =   15592941
      HoverColor      =   12632256
      ForeColor       =   8422784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "..."
      Align           =   0
   End
   Begin Emerald_Studio.EButton EButton1 
      Height          =   252
      Left            =   8088
      TabIndex        =   11
      Tag             =   "button:default"
      Top             =   2616
      Width           =   372
      _ExtentX        =   677
      _ExtentY        =   445
      DefaultColor    =   15592941
      HoverColor      =   12632256
      ForeColor       =   8422784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "..."
      Align           =   0
   End
   Begin Emerald_Studio.EButton EButton2 
      Height          =   252
      Left            =   8088
      TabIndex        =   14
      Tag             =   "button:default"
      Top             =   2976
      Width           =   372
      _ExtentX        =   677
      _ExtentY        =   445
      DefaultColor    =   15592941
      HoverColor      =   12632256
      ForeColor       =   8422784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "..."
      Align           =   0
   End
   Begin Emerald_Studio.EButton EButton3 
      Height          =   252
      Left            =   8616
      TabIndex        =   16
      Tag             =   "button:default"
      Top             =   2976
      Width           =   1068
      _ExtentX        =   1884
      _ExtentY        =   445
      DefaultColor    =   15592941
      HoverColor      =   12632256
      ForeColor       =   8422784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "获取SDK"
      Align           =   0
   End
   Begin Emerald_Studio.ECheckBox forcheck 
      Height          =   228
      Index           =   0
      Left            =   360
      TabIndex        =   26
      Tag             =   "switch"
      Top             =   3744
      Width           =   3372
      _ExtentX        =   5948
      _ExtentY        =   402
      OffColor        =   12632256
      OnColor         =   13556250
      Content         =   ".erp Emerald工程文件"
      IsOn            =   0   'False
      ForeColor       =   8422784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
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
      Left            =   360
      TabIndex        =   27
      Tag             =   "switch"
      Top             =   4080
      Width           =   3948
      _ExtentX        =   6964
      _ExtentY        =   402
      OffColor        =   12632256
      OnColor         =   13556250
      Content         =   ".ers Emerald Studio设置文件"
      IsOn            =   0   'False
      ForeColor       =   8422784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
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
      Index           =   2
      Left            =   360
      TabIndex        =   28
      Tag             =   "switch"
      Top             =   4416
      Width           =   3372
      _ExtentX        =   5948
      _ExtentY        =   402
      OffColor        =   12632256
      OnColor         =   13556250
      Content         =   ".erd Emerald存档文件"
      IsOn            =   0   'False
      ForeColor       =   8422784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label UIOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Infinity"
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
      Height          =   288
      Index           =   3
      Left            =   6288
      TabIndex        =   29
      Tag             =   "switch.off"
      Top             =   1848
      Width           =   1320
   End
   Begin VB.Label datawarning 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Caption         =   "您没有给予本程序在本地储存数据的权利，该页面的所有内容被禁用。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   288
      Left            =   -24
      TabIndex        =   25
      Top             =   5904
      Visible         =   0   'False
      Width           =   9936
   End
   Begin VB.Label switchpad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   24
      Tag             =   "switch.off"
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label switchpad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDA1A&
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
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   23
      Tag             =   "switch.on"
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "最大撤销步骤"
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
      Left            =   384
      TabIndex        =   21
      Tag             =   "text.title2"
      Top             =   5472
      Width           =   1176
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "工程自动保存间隔时间（分钟）"
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
      Left            =   384
      TabIndex        =   19
      Tag             =   "text.title2"
      Top             =   5136
      Width           =   2736
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "恢复"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   288
      Left            =   384
      TabIndex        =   18
      Tag             =   "text.title"
      Top             =   4800
      Width           =   396
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "关联"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   288
      Left            =   360
      TabIndex        =   17
      Tag             =   "text.title"
      Top             =   3360
      Width           =   396
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SDK 储存位置"
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
      Left            =   360
      TabIndex        =   13
      Tag             =   "text.title2"
      Top             =   2952
      Width           =   1200
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VB6 IDE 位置"
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
      Left            =   360
      TabIndex        =   10
      Tag             =   "text.title2"
      Top             =   2592
      Width           =   1176
   End
   Begin VB.Label UIOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Icelolly UI"
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
      Height          =   288
      Index           =   2
      Left            =   4776
      TabIndex        =   9
      Tag             =   "switch.off"
      Top             =   1848
      Width           =   1320
   End
   Begin VB.Label UIOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dark"
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
      Height          =   288
      Index           =   1
      Left            =   3264
      TabIndex        =   8
      Tag             =   "switch.off"
      Top             =   1848
      Width           =   1320
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "调试"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   288
      Left            =   360
      TabIndex        =   7
      Tag             =   "text.title"
      Top             =   2280
      Width           =   396
   End
   Begin VB.Label UIOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDA1A&
      Caption         =   "Light"
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
      Height          =   288
      Index           =   0
      Left            =   1728
      TabIndex        =   6
      Tag             =   "switch.on"
      Top             =   1848
      Width           =   1320
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "UI主题"
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
      Left            =   360
      TabIndex        =   5
      Tag             =   "text.title2"
      Top             =   1848
      Width           =   600
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "默认储存位置"
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
      Left            =   360
      TabIndex        =   2
      Tag             =   "text.title2"
      Top             =   1464
      Width           =   1176
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "常规"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   288
      Left            =   360
      TabIndex        =   1
      Tag             =   "text.title"
      Top             =   1104
      Width           =   396
   End
   Begin VB.Label title 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "设置"
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
      Left            =   360
      TabIndex        =   0
      Tag             =   "text.title"
      Top             =   312
      Width           =   396
   End
   Begin VB.Label background 
      BackColor       =   &H00F0F2F0&
      Height          =   924
      Left            =   0
      TabIndex        =   30
      Tag             =   "window.tool"
      Top             =   0
      Width           =   9900
   End
End
Attribute VB_Name = "SetWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    UpdateSkin Me, CurrentSkin
    
    '更新选项卡颜色
    For i = 0 To UIOption.UBound
        UIOption(i).Backcolor = switchpad(1).Backcolor
    Next
    UIOption(CurrentSkin).Backcolor = switchpad(0).Backcolor
    
    '如果没有给予存档权限，则禁用该页面所有功能
    If Data.sToken = False Then
        On Error Resume Next
        For Each co In Me.Controls
            co.Visible = False
        Next
        datawarning.Visible = True
        Me.Backcolor = datawarning.Backcolor
        datawarning.top = Me.ScaleHeight / 2 - datawarning.Height / 2
    End If
End Sub

Private Sub UIOption_Click(Index As Integer)
    '皮肤没有变化则退出
    If Index = CurrentSkin Then Exit Sub
    
    Dim f As Form
    '更新所有窗口的皮肤
    CurrentSkin = Index
    For Each f In VB.Forms
        UpdateSkin f, CurrentSkin
    Next
    '保存皮肤编号
    Data.PutData "theme", Index
    
    '更新选项卡颜色
    For i = 0 To UIOption.UBound
        UIOption(i).Backcolor = switchpad(1).Backcolor
    Next
    UIOption(Index).Backcolor = switchpad(0).Backcolor
End Sub
