Attribute VB_Name = "Core"
'=========================================================================
'   DPI适应
    Public Declare Function SetProcessDpiAwareness Lib "SHCORE.DLL" (ByVal DPImodel As Long) As Long
'=========================================================================
'   各种常量和结构体

'   是否在进行调试
    Public Const IsDebug As Boolean = False
    
'   设计窗口物件
    Public Type ESObj
        name As String                              '名称
        Content As String                           '内容
        style As Integer                            '样式
        size As Long                                '文本大小
        align As Integer                            '文本对齐方式
        actived As Boolean                          '内容是否属于活动性
        clicked As Boolean                          '该物件是否接受鼠标事件
        kind As Integer                             '类型
        Color As Long                               '使用颜色
        pad As Object                               '绑定的显示器
    End Type
    
'   设计窗口
    Public Type dsnWindows
        win As dsnWindow                            '对应的窗口
        Obj() As ESObj                              '该窗口的物件集合
    End Type
'========================================================================================
'   公有变量
    Public Data As GSaving                          '存档类、
    Public DW As Long, DH As Long                   '工程要求的窗口宽高
    Public DsnWin() As dsnWindows                   '设计窗口集合
    Public Droping As Boolean, DropI As Integer     '是否在拖放物件;拖放物件的目标类型
'========================================================================================
'   启动
    Public Sub Main()
    
        '初始化设计窗口集合
        ReDim DsnWin(0)
    
        If Val(GetWinNTVersion) > 6.1 Then               '如果当前系统版本高于win7
            SetProcessDpiAwareness 2&                    '调用API使本程序在高DPI情况下不模糊
        End If
        
        Set Data = New GSaving
        Data.Create "Emerald Studio", "Emerald Studio"   '创建存档
        CurrentSkin = Val(Data.GetData("theme"))         '从存档中取得当前主题编号
    
        '初始化Emerald
        StartEmerald MainWindow.hwnd, MainWindow.ScaleWidth, MainWindow.ScaleHeight
        '初始化字体
        MakeFont "微软雅黑"
        
        #If IsDebug Then                                 '是否在进行调试
            SkinTest.Show
            SkinTest.Left = 0
            ProjectWindow.Show
            ProjectWindow.Left = SkinTest.Width
        #Else
            StartupWindow.Show                           '主窗口，我们走！
        #End If
        
    End Sub
'========================================================================================
'   运行时
'   取得当前系统的WinNT版本
    Public Function GetWinNTVersion() As String
        Dim strComputer, objWMIService, colItems, objItem, strOSversion As String
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
        
        For Each objItem In colItems
            strOSversion = objItem.Version
        Next
        
        GetWinNTVersion = Left(strOSversion, 3)
    End Function
'   检查文件名是否非法
'   <name:文件名>
    Public Function CheckFileName(name As String) As Boolean
        CheckFileName = ((InStr(name, "*") Or InStr(name, "\") Or InStr(name, "/") Or InStr(name, ":") Or InStr(name, "?") Or InStr(name, """") Or InStr(name, "<") Or InStr(name, ">") Or InStr(name, "|")) = 0)
    End Function
'   添加新的设计窗口
'   <f:设计窗口>
    Public Sub AddDsnWindow(f As dsnWindow)
        ReDim Preserve DsnWin(UBound(DsnWin) + 1)
        Set DsnWin(UBound(DsnWin)).win = f
        ReDim DsnWin(UBound(DsnWin)).Obj(0)
    End Sub
'========================================================================================
