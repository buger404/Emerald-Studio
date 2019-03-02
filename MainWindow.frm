VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "image.ocx"
Begin VB.Form MainWindow 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Emerald Studio"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   768
   ClientWidth     =   15000
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1250
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame TipFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEDA1A&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8280
      Width           =   15015
      Begin ImageX.aicAlphaImage bmFace 
         Height          =   372
         Left            =   24
         Top             =   0
         Width           =   492
         _ExtentX        =   868
         _ExtentY        =   656
         Image           =   "MainWindow.frx":1BCC2
         Props           =   5
      End
      Begin VB.Label tiptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "一切准备就绪汪。"
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
         Left            =   576
         TabIndex        =   49
         Top             =   48
         Width           =   1536
      End
      Begin VB.Label frame_back 
         BackColor       =   &H00CEDA1A&
         Height          =   400
         Left            =   -48
         TabIndex        =   50
         Tag             =   "window.highlight"
         Top             =   -48
         Width           =   17012
      End
   End
   Begin VB.PictureBox proframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F2F0&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   0
      ScaleHeight     =   5052
      ScaleWidth      =   3612
      TabIndex        =   7
      Tag             =   "window.tool"
      Top             =   3240
      Width           =   3615
      Begin VB.PictureBox proframe_h 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F2F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Left            =   96
         ScaleHeight     =   3732
         ScaleWidth      =   3252
         TabIndex        =   41
         Tag             =   "window.tool"
         Top             =   4848
         Width           =   3255
      End
      Begin VB.PictureBox fontcover 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F2F0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1236
         Left            =   3408
         ScaleHeight     =   1236
         ScaleWidth      =   3252
         TabIndex        =   58
         Tag             =   "window.tool"
         Top             =   3720
         Width           =   3255
      End
      Begin Emerald_Studio.FCombo objCombo 
         Height          =   324
         Left            =   240
         TabIndex        =   53
         Tag             =   "combo:tool"
         Top             =   672
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   572
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   10.2
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Backcolor       =   15790832
         Forecolor       =   8421504
         Fontcolor       =   8421504
         Caption         =   ""
      End
      Begin Emerald_Studio.EEdit protext 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   33
         Tag             =   "edit:tool"
         ToolTipText     =   "元素的内容"
         Top             =   1584
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   15790832
         Content         =   "test"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit protext 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   32
         Tag             =   "edit:tool"
         ToolTipText     =   "元素的名称，可以为空"
         Top             =   1224
         Width           =   1716
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   15790832
         Content         =   "test"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit protext 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   31
         Tag             =   "edit:tool"
         ToolTipText     =   "字体大小，只有在文字元素中有效"
         Top             =   4176
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   15790832
         Content         =   "16"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit protext 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   29
         Tag             =   "edit:tool"
         ToolTipText     =   "字体样式，只有在文字元素中有效"
         Top             =   3816
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   15790832
         Content         =   "0 - Regular"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit protext 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   27
         Tag             =   "edit:tool"
         ToolTipText     =   "元素高度，在图形中禁用"
         Top             =   3288
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   15790832
         Content         =   "0"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit protext 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   25
         Tag             =   "edit:tool"
         ToolTipText     =   "元素宽度，在图形中禁用"
         Top             =   2928
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   15790832
         Content         =   "0"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit protext 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   23
         Tag             =   "edit:tool"
         ToolTipText     =   "Y坐标"
         Top             =   2568
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   15790832
         Content         =   "0"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit protext 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   21
         Tag             =   "edit:tool"
         ToolTipText     =   "X坐标"
         Top             =   2208
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   15790832
         Content         =   "0"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin VB.Label alignflag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "C"
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
         Left            =   3024
         TabIndex        =   57
         Tag             =   "switch.off"
         ToolTipText     =   "鼠标检测标记"
         Top             =   1224
         Width           =   252
      End
      Begin VB.Label alignflag 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "R"
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
         Left            =   2640
         TabIndex        =   40
         Tag             =   "switch.off"
         ToolTipText     =   "向右对齐"
         Top             =   4536
         Width           =   492
      End
      Begin VB.Label alignflag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "C"
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
         Left            =   1920
         TabIndex        =   39
         Tag             =   "switch.off"
         ToolTipText     =   "居中对齐"
         Top             =   4536
         Width           =   492
      End
      Begin VB.Label alignflag 
         Appearance      =   0  'Flat
         BackColor       =   &H00CEDA1A&
         Caption         =   "L"
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
         Left            =   1200
         TabIndex        =   38
         Tag             =   "switch.on"
         ToolTipText     =   "向左对齐"
         Top             =   4536
         Width           =   492
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Align"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9.6
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   288
         Index           =   8
         Left            =   240
         TabIndex        =   37
         Tag             =   "text.title2"
         Top             =   4536
         Width           =   468
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Content"
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
         Index           =   4
         Left            =   240
         TabIndex        =   36
         Tag             =   "text.title2"
         Top             =   1584
         Width           =   756
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Index           =   7
         Left            =   240
         TabIndex        =   35
         Tag             =   "text.title2"
         Top             =   1224
         Width           =   552
      End
      Begin VB.Label alignflag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "A"
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
         Index           =   4
         Left            =   3024
         TabIndex        =   34
         Tag             =   "switch.off"
         ToolTipText     =   "活动元素标记"
         Top             =   1560
         Width           =   252
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
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
         Index           =   6
         Left            =   240
         TabIndex        =   30
         Tag             =   "text.title2"
         Top             =   4176
         Width           =   360
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Style"
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
         Index           =   5
         Left            =   240
         TabIndex        =   28
         Tag             =   "text.title2"
         Top             =   3816
         Width           =   456
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         Tag             =   "line"
         X1              =   240
         X2              =   3240
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Tag             =   "line"
         X1              =   240
         X2              =   3240
         Y1              =   2016
         Y2              =   2016
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
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
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Tag             =   "text.title2"
         Top             =   3288
         Width           =   612
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
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
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Tag             =   "text.title2"
         Top             =   2928
         Width           =   552
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
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
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Tag             =   "text.title2"
         Top             =   2568
         Width           =   120
      End
      Begin VB.Label ptext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Tag             =   "text.title2"
         Top             =   2208
         Width           =   120
      End
      Begin VB.Label protitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "属性列表"
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
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Tag             =   "text.title2"
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.PictureBox colorframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F2F0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   11400
      ScaleHeight     =   4812
      ScaleWidth      =   3612
      TabIndex        =   5
      Tag             =   "window.tool"
      Top             =   0
      Width           =   3615
      Begin Emerald_Studio.EEdit colortext 
         Height          =   300
         Index           =   4
         Left            =   2928
         TabIndex        =   42
         Tag             =   "edit:tool"
         ToolTipText     =   "颜色的Alpha值，需要控制在0~255之间。"
         Top             =   3120
         Width           =   516
         _ExtentX        =   910
         _ExtentY        =   529
         BackColor       =   15790832
         Content         =   "242"
         ForeColor       =   8421504
         BorderColor     =   13556506
         Alignment       =   2
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit colortext 
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   19
         Tag             =   "edit:tool"
         ToolTipText     =   "HEX颜色，你可以粘贴一个新的HEX颜色代码或复制它。"
         Top             =   4440
         Width           =   876
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   15790832
         Content         =   "DEDEDE"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   2
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit colortext 
         Height          =   300
         Index           =   2
         Left            =   1200
         TabIndex        =   17
         Tag             =   "edit:tool"
         ToolTipText     =   "颜色的Blue值，需要控制在0~255之间。"
         Top             =   3960
         Width           =   1044
         _ExtentX        =   1842
         _ExtentY        =   529
         BackColor       =   15790832
         Content         =   "242"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit colortext 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Tag             =   "edit:tool"
         ToolTipText     =   "颜色的Green值，需要控制在0~255之间。"
         Top             =   3600
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         BackColor       =   15790832
         Content         =   "242"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin Emerald_Studio.EEdit colortext 
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   15
         Tag             =   "edit:tool"
         ToolTipText     =   "颜色的Red值，需要控制在0~255之间。"
         Top             =   3240
         Width           =   996
         _ExtentX        =   1757
         _ExtentY        =   529
         BackColor       =   15790832
         Content         =   "242"
         ForeColor       =   9871510
         BorderColor     =   13556506
         Alignment       =   0
         LockInput       =   0   'False
      End
      Begin VB.PictureBox colorpad 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DEDEDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   2
         Left            =   360
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   241
         TabIndex        =   11
         ToolTipText     =   "各种颜色调整板"
         Top             =   2640
         Width           =   2895
         Begin VB.Shape colorpoint 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   2
            Left            =   0
            Shape           =   2  'Oval
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.PictureBox colorpad 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DEDEDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   360
         ScaleHeight     =   11
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   241
         TabIndex        =   10
         ToolTipText     =   "颜色透明度调整"
         Top             =   2880
         Width           =   2895
         Begin VB.Shape colorpoint 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   1
            Left            =   1320
            Shape           =   2  'Oval
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.PictureBox colorpad 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00DEDEDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   0
         Left            =   360
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   9
         ToolTipText     =   "颜色调整面板"
         Top             =   720
         Width           =   1695
         Begin VB.Shape colorpoint 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   135
            Index           =   0
            Left            =   840
            Shape           =   2  'Oval
            Top             =   720
            Width           =   135
         End
      End
      Begin Emerald_Studio.EButton applycolorbutton 
         Height          =   372
         Left            =   2328
         TabIndex        =   59
         Tag             =   "button:tool"
         Top             =   4392
         Width           =   1092
         _ExtentX        =   1926
         _ExtentY        =   656
         DefaultColor    =   15790832
         HoverColor      =   15592941
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9.6
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Content         =   "应用此颜色"
         Align           =   0
      End
      Begin ImageX.aicAlphaImage dropper 
         Height          =   288
         Left            =   2952
         Top             =   2112
         Width           =   288
         _ExtentX        =   508
         _ExtentY        =   508
         Image           =   "MainWindow.frx":37300
         Props           =   5
      End
      Begin VB.Label colormem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   492
         Index           =   0
         Left            =   2400
         TabIndex        =   43
         Top             =   1056
         Width           =   492
      End
      Begin VB.Label colormem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   492
         Index           =   1
         Left            =   2640
         TabIndex        =   44
         Top             =   1296
         Width           =   492
      End
      Begin VB.Label ctext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Code   #"
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
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Tag             =   "text.title2"
         Top             =   4440
         Width           =   780
      End
      Begin VB.Label ctext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
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
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Tag             =   "text.title2"
         Top             =   3960
         Width           =   390
      End
      Begin VB.Label ctext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
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
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Tag             =   "text.title2"
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label ctext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
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
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Tag             =   "text.title2"
         Top             =   3240
         Width           =   345
      End
      Begin VB.Label colortitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "调色板"
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
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Tag             =   "text.title2"
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.PictureBox expframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F2F0&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3252
      ScaleWidth      =   3612
      TabIndex        =   3
      Tag             =   "window.tool"
      Top             =   0
      Width           =   3615
      Begin VB.Label exptitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "资源管理器"
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
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Tag             =   "text.title2"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox ToolFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F2F0&
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   11400
      ScaleHeight     =   3492
      ScaleWidth      =   3612
      TabIndex        =   0
      Tag             =   "window.tool"
      Top             =   4800
      Width           =   3615
      Begin Emerald_Studio.EButton toolitems 
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Tag             =   "button:tool"
         Top             =   600
         Width           =   3135
         _ExtentX        =   5525
         _ExtentY        =   868
         DefaultColor    =   15790832
         HoverColor      =   15592941
         ForeColor       =   9871510
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9.6
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Content         =   "图形元素"
         Align           =   1
      End
      Begin Emerald_Studio.EButton toolitems 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   47
         Tag             =   "button:tool"
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5525
         _ExtentY        =   868
         DefaultColor    =   15790832
         HoverColor      =   15592941
         ForeColor       =   9871510
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9.6
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Content         =   "文字元素"
         Align           =   1
      End
      Begin Emerald_Studio.EButton toolitems 
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   48
         Tag             =   "button:tool"
         Top             =   1560
         Width           =   3135
         _ExtentX        =   5525
         _ExtentY        =   868
         DefaultColor    =   15790832
         HoverColor      =   15592941
         ForeColor       =   9871510
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9.6
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Content         =   "形状元素"
         Align           =   1
      End
      Begin VB.Label tooltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "工具箱"
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
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Tag             =   "text.title2"
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.PictureBox designf 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   3624
      ScaleHeight     =   8292
      ScaleWidth      =   7812
      TabIndex        =   45
      Tag             =   "window.background"
      Top             =   0
      Width           =   7815
      Begin VB.Timer tiptimer 
         Interval        =   100
         Left            =   7200
         Top             =   264
      End
      Begin VB.PictureBox toolframe2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F2F0&
         BorderStyle     =   0  'None
         Height          =   708
         Left            =   840
         ScaleHeight     =   708
         ScaleWidth      =   6228
         TabIndex        =   52
         Tag             =   "window.tool"
         Top             =   7584
         Width           =   6228
         Begin Emerald_Studio.FCombo pagelist 
            Height          =   348
            Left            =   240
            TabIndex        =   54
            Tag             =   "combo:tool"
            Top             =   192
            Width           =   3108
            _ExtentX        =   5482
            _ExtentY        =   614
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "微软雅黑"
               Size            =   10.2
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Backcolor       =   15790832
            Forecolor       =   8421504
            Fontcolor       =   8421504
            Caption         =   ""
         End
      End
      Begin VB.Frame dsnpane 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4500
         Left            =   1632
         TabIndex        =   51
         Top             =   1584
         Width           =   4668
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
         Height          =   288
         Index           =   0
         Left            =   216
         TabIndex        =   56
         Tag             =   "switch.on"
         Top             =   192
         Visible         =   0   'False
         Width           =   240
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
         Height          =   288
         Index           =   1
         Left            =   576
         TabIndex        =   55
         Tag             =   "switch.off"
         Top             =   192
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Menu 文件 
      Caption         =   "文件(&F)"
      Begin VB.Menu opencmd 
         Caption         =   "打开(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu newcmd 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu splitline0 
         Caption         =   "-"
      End
      Begin VB.Menu savecmd 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu splitline1 
         Caption         =   "-"
      End
      Begin VB.Menu closecmd 
         Caption         =   "关闭(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu editcmd 
      Caption         =   "编辑(&E)"
      Begin VB.Menu copycmd 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu pastecmd 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu splitline2 
         Caption         =   "-"
      End
      Begin VB.Menu delcmd 
         Caption         =   "移除(&D)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu runcmd 
      Caption         =   "生成(&R)"
      Begin VB.Menu openvbcmd 
         Caption         =   "打开 Visual Studio 6.0 工程"
      End
      Begin VB.Menu excutecmd 
         Caption         =   "生成可执行文件"
      End
      Begin VB.Menu packcmd 
         Caption         =   "一键打包安装程序"
      End
   End
   Begin VB.Menu toolcmd 
      Caption         =   "工具(&T)"
      Begin VB.Menu setcmd 
         Caption         =   "设置"
      End
      Begin VB.Menu colorgetcmd 
         Caption         =   "配色采集器"
      End
      Begin VB.Menu designiconcmd 
         Caption         =   "图形设计器"
      End
      Begin VB.Menu prosetcmd 
         Caption         =   "工程设置"
      End
   End
   Begin VB.Menu helpcmd 
      Caption         =   "帮助(&H)"
      Begin VB.Menu helpdoccmd 
         Caption         =   "帮助文档"
      End
      Begin VB.Menu splitline4 
         Caption         =   "-"
      End
      Begin VB.Menu aboutcmd 
         Caption         =   "关于我"
      End
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pad_a As Long, pad_r As Long, _
    pad_g As Long, pad_b As Long, _
    keyc As Boolean                     '调色板颜色（keyc：是否是因为用户的输入而改变颜色）
Dim lTip As Object                      '当前的已提示的对象
Public nPage As Integer                 '当前的设计窗口
'新的提示信息
'<tips：提示内容,face:提示图像>
Public Sub NewTip(ByVal tips As String, Optional face As String = "bm_ok.png")
    '更新黑嘴的提示内容
    bmFace.LoadImage_FromFile App.path & "\assets\" & face
    tiptext.Caption = tips
End Sub
Private Sub aboutcmd_Click()
    AboutWindow.Show
End Sub

Private Sub alignflag_Click(Index As Integer)
    With DsnWin(nPage).Obj(objCombo.ListIndex + 1)
        If Index = 3 Then .clicked = Not .clicked: alignflag(Index).Backcolor = switchpad(IIf(.clicked, 0, 1)).Backcolor: Exit Sub
        If Index = 4 Then .actived = Not .actived: alignflag(Index).Backcolor = switchpad(IIf(.actived, 0, 1)).Backcolor: Exit Sub
        For i = 0 To 2
            alignflag(i).Backcolor = switchpad(1).Backcolor
        Next
        alignflag(Index).Backcolor = switchpad(0).Backcolor
        .align = Index
        DsnWin(nPage).win.UpdateUS objCombo.ListIndex + 1
    End With
End Sub

Private Sub applycolorbutton_Click()
    '判断当前是否选中了控件
    If objCombo.ListIndex = -1 Then Exit Sub

    With DsnWin(nPage).Obj(objCombo.ListIndex + 1)
        '判断是否支持颜色
        If .kind <> 0 Then
            .Color = argb(pad_a, pad_r, pad_g, pad_b)
        End If
    End With
    
    '刷新控件
    DsnWin(nPage).win.UpdateUS objCombo.ListIndex + 1
End Sub

Private Sub bmFace_Click(ByVal Button As Integer)
    NewTip "不要打我o(>﹏<)o", "bm_sad.png"
End Sub
Private Sub closecmd_Click()
    Unload Me
End Sub
Private Sub colorpad_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Dim c(3) As Byte
        Select Case Index
        Case 0                                                                  '大调色板
            If X < 0 Then X = 0
            If X > colorpad(0).ScaleWidth - 1 Then X = colorpad(0).ScaleWidth - 1
            If Y < 0 Then Y = 0
            If Y > colorpad(0).ScaleHeight - 1 Then Y = colorpad(0).ScaleHeight - 1
            CopyMemory c(0), colorpad(0).POINT(X, Y), 4
            pad_r = c(0): pad_g = c(1): pad_b = c(2)
            Call setCPadC
            Call UpdatePadPP
            colorpoint(0).Move X - 4.5, Y - 4.5
        Case 1                                                                  '透明度
            If X < 0 Then X = 0
            If X > colorpad(1).ScaleWidth Then X = colorpad(1).ScaleWidth
            pad_a = X / colorpad(1).ScaleWidth * 255
            Call setCPadC
            colorpoint(1).Move X - 4.5
        Case 2                                                                  '小调色板
            If X < 0 Then X = 0
            If X > colorpad(2).ScaleWidth - 1 Then X = colorpad(2).ScaleWidth - 1
            CopyMemory c(0), colorpad(2).POINT(X, 0), 4
            pad_r = c(0): pad_g = c(1): pad_b = c(2)
            Call setCPadC
            Call setCPad
            Call UpdatePadPP
            colorpoint(2).Move X - 4.5
        End Select
    End If
End Sub

Private Sub colorpad_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call colorpad_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub colortext_Change(Index As Integer)
    If keyc = False Then Exit Sub
    
    keyc = False
    If Index = 3 Then
        Dim c(3) As Byte, l As Long
        l = CLng("&H" & colortext(3).Content)
        CopyMemory c(0), l, 4
        'vb的Hex颜色和大众的不统一，统一一下
        pad_r = c(2): pad_g = c(1): pad_b = c(0)
        Call setCPadC
        Call setCPad
        Call UpdatePadPP
        Exit Sub
    End If
    
    If Val(colortext(Index).Content) > 255 Then colortext(Index).Content = Right(colortext(Index).Content, 2)
    If Val(colortext(Index).Content) < 0 Then colortext(Index).Content = 0
    
    pad_r = Val(colortext(0).Content): pad_g = Val(colortext(1).Content): pad_b = Val(colortext(2).Content)
    pad_a = Val(colortext(4).Content)
    Call setCPadC
    Call setCPad
    Call UpdatePadPP
End Sub

Private Sub colortext_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    keyc = True
End Sub

Private Sub dropper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = MousePointerConstants.vbCrosshair
End Sub

Private Sub dropper_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 0
    
    Dim p As POINT, dc As Long, co As Long, c(3) As Byte
    GetCursorPos p
    
    dc = GetDC(0)
    co = GetPixel(dc, p.X, p.Y)
    
    CopyMemory c(0), co, 4
    pad_a = 255: pad_r = c(0): pad_g = c(1): pad_b = c(2)
    Call setCPadC
    Call setCPad
    Call UpdatePadPP
End Sub

Private Sub Form_Load()
    proframe_h.Visible = True
    
    '调色板初始化
    pad_a = 255: pad_r = 255: pad_g = 0: pad_b = 0
    Call setCPadC
    Call setCPad
    Call UpdatePadPP
    
    '设计窗口初始化
    Dim test As New dsnWindow
    test.Show
    SetParent test.hwnd, dsnpane.hwnd                                           ': test.Move 0, 0
    pagelist.ListIndex = 0: nPage = 1
    
    '样式初始化
    UpdateSkin Me, CurrentSkin
    
    '窗口大小读取
    If Data.GetData("win_width") <> "" Then
        Me.Move Val(Data.GetData("win_left")), Val(Data.GetData("win_top")), Val(Data.GetData("win_width")), Val(Data.GetData("win_height"))
    End If
    
    '设置遮盖板的位置
    proframe_h.Move ptext(7).Left - 5, ptext(7).top - 5
    fontcover.Move ptext(5).Left - 5, ptext(5).top - 5
End Sub
Public Sub UpdatePane()
    '更新设计窗口坐标
    Dim w As Long, h As Long
    w = DW * Screen.TwipsPerPixelX: h = DH * Screen.TwipsPerPixelY
    dsnpane.Move designf.Width * Screen.TwipsPerPixelX / 2 - w / 2, designf.Height * Screen.TwipsPerPixelY / 2 - h / 2, w, h
End Sub
Public Sub UpdatePadPP()
    '更新调色板取色点坐标
    colorpoint(0).Move colorpad(0).ScaleWidth / 2 - 4.5, colorpad(0).ScaleHeight / 2 - 4.5
    colorpoint(1).Move pad_a / 255 * (colorpad(1).ScaleWidth - 4.5), 0
End Sub
Public Sub setCPadC()
    '更新颜色到文本框
    
    If colortext(0).Content <> pad_r Then colortext(0).Content = pad_r
    If colortext(1).Content <> pad_g Then colortext(1).Content = pad_g
    If colortext(2).Content <> pad_b Then colortext(2).Content = pad_b
    Dim strc As String
    strc = Hex(RGB(pad_b, pad_g, pad_r))
    For i = 1 To 6 - Len(strc)
        strc = "0" & strc
    Next
    If colortext(3).Content <> strc Then colortext(3).Content = strc
    
    If colortext(4).Content <> pad_a Then colortext(4).Content = pad_a
    colormem(0).Backcolor = RGB(pad_r, pad_g, pad_b)
End Sub
Public Sub setCPad()
    '设置大调色板颜色
    
    Dim r As Long, g As Long, b As Long
    r = pad_r: g = pad_g: b = pad_b
    
    Dim gr As Long, br As Long, p As Long
    Dim c() As Long, best As Long
    GdipCreateFromHDC colorpad(0).hdc, gr
    
    GdipCreatePath FillModeWinding, p
    GdipAddPathRectangle p, 0, 0, colorpad(0).ScaleWidth, colorpad(0).ScaleHeight
    
    ReDim c(3)
    c(0) = argb(255, 255, 255, 255): c(2) = argb(255, 0, 0, 0): c(3) = argb(255, 0, 0, 0)
    '计算饱和颜色
    best = IIf(r > g, r, g): best = IIf(best < b, b, best)
    If best = r Then c(1) = argb(255, 255, g, b)
    If best = g Then c(1) = argb(255, r, 255, b)
    If best = b Then c(1) = argb(255, r, g, 255)
    
    GdipCreatePathGradientFromPath p, br
    GdipSetPathGradientSurroundColorsWithCount br, c(0), UBound(c) + 1
    GdipSetPathGradientCenterColor br, argb(255, r, g, b)
    
    GdipFillPath gr, br, p
    
    GdipDeleteGraphics gr
    GdipDeleteBrush br
    GdipDeletePath p
    
    colorpad(0).Refresh
    
    Dim im As Long
    GdipCreateBitmapFromFile StrPtr(App.path & "\assets\alpha.png"), im
    
    GdipCreateFromHDC colorpad(1).hdc, gr
    GdipCreateLineBrush NewPointF(0, 0), NewPointF(colorpad(1).ScaleWidth, 0), argb(0, 255, 255, 255), argb(255, r, g, b), WrapModeTile, br
    GdipDrawImage gr, im, 0, 0
    GdipFillRectangle gr, br, 0, 0, colorpad(1).ScaleWidth, colorpad(1).ScaleHeight
    
    colorpad(1).Refresh
    
    GdipDeleteGraphics gr
    GdipDeleteBrush br
    GdipDisposeImage im
    
    GdipCreateFromHDC colorpad(2).hdc, gr
    GdipGraphicsClear gr, 0
    
    Dim a() As Single, i As Integer
    
    ReDim c(5), a(5)
    c(0) = argb(255, 255, 0, 0): c(1) = argb(255, 255, 255, 0): c(2) = argb(255, 0, 255, 0)
    c(3) = argb(255, 0, 255, 255): c(4) = argb(255, 0, 0, 255): c(5) = argb(255, 255, 0, 255)
    
    For i = 0 To UBound(a)
        a(i) = Int(i / 5 * 100) / 100
    Next
    
    GdipCreateLineBrush NewPointF(0, 0), NewPointF(colorpad(2).ScaleWidth - 1, 0), 0, 0, WrapModeTileFlipXY, br
    GdipSetLinePresetBlend br, c(0), a(0), UBound(c) + 1
    
    GdipFillRectangle gr, br, 0, 0, colorpad(2).ScaleWidth, colorpad(2).ScaleHeight
    
    GdipDeleteGraphics gr
    GdipDeleteBrush br
    
    colorpad(2).Refresh
End Sub

Private Sub Form_Resize()
    '更新各种控件的坐标
    
    On Error Resume Next
    expframe.Height = Me.ScaleHeight - TipFrame.Height - proframe.Height
    proframe.top = expframe.top + expframe.Height
    
    ToolFrame.Height = Me.ScaleHeight - TipFrame.Height - colorframe.Height
    ToolFrame.top = colorframe.Height
    ToolFrame.Left = Me.ScaleWidth - ToolFrame.Width
    colorframe.Left = ToolFrame.Left
    TipFrame.top = Me.ScaleHeight - TipFrame.Height
    TipFrame.Width = Me.ScaleWidth
    frame_back.Move 0, 0, TipFrame.Width * Screen.TwipsPerPixelX, TipFrame.Height * Screen.TwipsPerPixelY
    
    designf.Move expframe.Width, 0, Me.ScaleWidth - expframe.Width * 2, Me.ScaleHeight - TipFrame.Height
    
    toolframe2.top = 0
    toolframe2.Left = designf.Width * Screen.TwipsPerPixelX / 2 - toolframe2.Width / 2
    
    Call UpdatePane
    
    Data.PutData "win_width", Me.Width
    Data.PutData "win_height", Me.Height
    Data.PutData "win_left", Me.Left
    Data.PutData "win_top", Me.top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndEmerald
    On Error Resume Next
    Unload ProjectWindow
    Unload AboutWindow
    Unload CreateWindow
    Unload SetWindow
    Unload WarnWindow
    End
End Sub

Private Sub objCombo_Click()
    If objCombo.ListIndex = -1 Then Exit Sub
    proframe_h.Visible = False
    '设置设计窗口的活动物件
    DsnWin(nPage).win.SetFocusIn objCombo.ListIndex + 1
    
    '把物件的属性内容展示到文本框内
    With DsnWin(nPage).Obj(objCombo.ListIndex + 1)
        protext(7).Content = .name
        protext(4).Content = .Content
        protext(0).Content = .pad.Left
        protext(1).Content = .pad.top - DsnWin(nPage).win.background.Height
        protext(2).Content = .pad.Width
        protext(3).Content = .pad.Height
        fontcover.Visible = True
        If .kind = 1 Then
            fontcover.Visible = False
            Select Case .style
                Case 0
                    protext(5).Content = "0 - Regular"
                Case 1
                    protext(5).Content = "1 - Bold"
                Case 2
                    protext(5).Content = "2 - Italic"
                Case Else
                    protext(5).Content = "[err] - Unknown"
            End Select
            protext(6).Content = .size
        End If
        
        For i = 0 To alignflag.UBound
            alignflag(i).Backcolor = switchpad(1).Backcolor
        Next
        alignflag(.align).Backcolor = switchpad(0).Backcolor
        If .clicked Then alignflag(3).Backcolor = switchpad(0).Backcolor
        If .actived Then alignflag(4).Backcolor = switchpad(0).Backcolor
    End With
End Sub

Private Sub pagelist_Click()
    If pagelist.ListIndex <> nPage - 1 Then
        nPage = pagelist.ListIndex + 1
        objCombo.Clear
        proframe_h.Visible = True
        '添加设计窗口的物件到物件列表
        For i = 1 To UBound(DsnWin(pagelist.ListIndex + 1).Obj)
            '是否拥有自己的名字
            If DsnWin(pagelist.ListIndex + 1).Obj(i).name = "" Then
                objCombo.AddItem "[Object " & i & " ]"
            Else
                objCombo.AddItem DsnWin(pagelist.ListIndex + 1).Obj(i).name
            End If
        Next
        '如果该设计窗口没有任何物件
        If UBound(DsnWin(pagelist.ListIndex + 1).Obj) > 0 Then
            '隐藏属性表
            proframe_h.Visible = False
            objCombo.ListIndex = 0
        End If
    End If
End Sub

Private Sub protext_Commit(Index As Integer)
    '<index[0=x,1=y,2=width,3=height,4=content,5=style,6=size,7=name]>
    Dim i As Integer
    i = objCombo.ListIndex + 1
    
    '文本格式处理
    If Index <> 4 And Index <> 7 And Index <> 5 Then
        protext(Index).Content = Val(protext(Index).Content)
    End If
    '特殊格式
    If DsnWin(nPage).Obj(i).kind = 2 And Index = 4 Then protext(Index).Content = Val(protext(Index).Content)
    If Index = 5 Then protext(Index).Content = Val(Left(protext(Index).Content, 1))
    
    Select Case Index
        Case 0: DsnWin(nPage).win.us(i).Left = Val(protext(Index).Content)
        Case 1: DsnWin(nPage).win.us(i).top = Val(protext(Index).Content) + DsnWin(nPage).win.titleframe.Height
        Case 2: DsnWin(nPage).win.us(i).Width = Val(protext(Index).Content)
        Case 3: DsnWin(nPage).win.us(i).Height = Val(protext(Index).Content)
        Case 4: DsnWin(nPage).Obj(i).Content = protext(Index).Content
        Case 5: DsnWin(nPage).Obj(i).style = Val(protext(Index).Content)
        Case 6: DsnWin(nPage).Obj(i).size = Val(protext(Index).Content)
        Case 7: DsnWin(nPage).Obj(i).name = protext(Index).Content: objCombo.Text = IIf(protext(Index).Content = "", "[Object " & i & "]", protext(Index).Content): objCombo.List(i - 1) = objCombo.Text: Exit Sub
    End Select
    
    '和坐标大小有关
    If Index <= 3 Then
        DsnWin(nPage).win.SetFocusIn i
        If Index <= 1 Then Exit Sub
    End If
    
    If Index = 5 Then
        Select Case Val(protext(Index).Content)
            Case 0
                protext(5).Content = "0 - Regular"
            Case 1
                protext(5).Content = "1 - Bold"
            Case 2
                protext(5).Content = "2 - Italic"
            Case Else
                protext(5).Content = "[err] - Unknown"
        End Select
    End If
    
    '更新
    DsnWin(nPage).win.UpdateUS i
End Sub

Private Sub setcmd_Click()
    SetWindow.Show
End Sub

Private Sub tiptimer_Timer()
    On Error Resume Next
    '如果当前活动控件未进行提示
    If Not (lTip Is Me.ActiveControl) Then
        Set lTip = Me.ActiveControl
        If Me.ActiveControl.ToolTipText = "" Then
            NewTip "没有可用的提示信息"
        Else
            NewTip lTip.ToolTipText
        End If
    End If
End Sub

Private Sub toolitems_Click(Index As Integer)
    If Screen.MousePointer = 0 Then
        Screen.MousePointer = MousePointerConstants.vbCrosshair
        Droping = True: DropI = Index           '开始拽拖
    Else
        Screen.MousePointer = 0
        Droping = False
    End If
End Sub
