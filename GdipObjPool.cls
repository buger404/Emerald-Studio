VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GdipObjPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'用于创建和销毁Gdip对象
Option Explicit
Private Token As Long


Dim Pens() As Long, PenCount As Long
Dim Brushes() As Long, BrushCount As Long
Dim StrFormats() As Long, StrFormatCount As Long
Dim Matrixes() As Long, MatrixCount As Long

Private Sub Class_Initialize()
        Dim uInput As GdiplusStartupInput
        Dim ret    As GpStatus
        
        uInput.GdiplusVersion = 1
        ret = GdiplusStartup(Token, uInput)
End Sub

Private Function DeleteObjects()
    Dim i As Long

    For i = 1 To PenCount: GdipDeletePen Pens(i): Next
    For i = 1 To BrushCount: GdipDeleteBrush Brushes(i): Next
    For i = 1 To StrFormatCount: GdipDeleteStringFormat StrFormats(i): Next
    For i = 1 To MatrixCount: GdipDeleteMatrix Matrixes(i): Next
    PenCount = 0
    BrushCount = 0
    StrFormatCount = 0
    MatrixCount = 0
End Function

Public Function NewPen(ByVal Color As Long, ByVal Width As Single) As Long
    PenCount = PenCount + 1
    ReDim Preserve Pens(PenCount)
    
    GdipCreatePen1 Color, Width, UnitPixel, Pens(PenCount)
    NewPen = Pens(PenCount)
End Function

Public Function NewBrush(ByVal Color As Long) As Long
    BrushCount = BrushCount + 1
    ReDim Preserve Brushes(BrushCount)
    
    GdipCreateSolidFill Color, Brushes(BrushCount)
    NewBrush = Brushes(BrushCount)
End Function

Public Function NewStringFormat(ByVal align As StringAlignment) As Long
    StrFormatCount = StrFormatCount + 1
    ReDim Preserve StrFormats(StrFormatCount)
    
    GdipCreateStringFormat 0, 0, StrFormats(StrFormatCount)
    GdipSetStringFormatAlign StrFormats(StrFormatCount), align
    NewStringFormat = StrFormats(StrFormatCount)
End Function

Public Function NewMatrix(ByVal m11 As Single, _
                          ByVal m12 As Single, _
                          ByVal m21 As Single, _
                          ByVal m22 As Single, _
                          ByVal DX As Single, _
                          ByVal DY As Single) As Long

    MatrixCount = MatrixCount + 1
    ReDim Preserve Matrixes(MatrixCount)
    
    GdipCreateMatrix Matrixes(MatrixCount)
    GdipSetMatrixElements Matrixes(MatrixCount), m11, m12, m21, m22, DX, DY
    NewMatrix = Matrixes(MatrixCount)
End Function




Public Function NewPointsFPtr(ParamArray ptXY()) As Long
    If (UBound(ptXY) And 1) = 0 Then GoTo ErrHandle
    
    Dim ret() As POINTF, i As Long
    ReDim ret(0 To UBound(ptXY) \ 2)
    
    For i = 0 To UBound(ptXY) Step 2
        ret(i \ 2).X = ptXY(i)
        ret(i \ 2).Y = ptXY(i + 1)
    Next
    
    NewPointsFPtr = VarPtr(ret(0))
    
    Exit Function
ErrHandle:
    NewPointsFPtr = 0
End Function

Public Function NewPointsLPtr(ParamArray ptXY()) As Long
    If (UBound(ptXY) And 1) = 0 Then GoTo ErrHandle
    
    Dim ret() As POINTL, i As Long
    ReDim ret(0 To UBound(ptXY) \ 2)
    
    For i = 0 To UBound(ptXY) Step 2
        ret(i \ 2).X = ptXY(i)
        ret(i \ 2).Y = ptXY(i + 1)
    Next
    
    NewPointsLPtr = VarPtr(ret(0))
    
    Exit Function
ErrHandle:
    NewPointsLPtr = 0
End Function

Public Function NewColors(ParamArray colors()) As Long()
    Dim ret() As Long, i As Long
    
    ReDim ret(UBound(colors))
    For i = 0 To UBound(colors)
        ret(i) = colors(i)
    Next
    
    NewColors = ret
End Function

Public Sub Dispose() '销毁所有对象
    Call DeleteObjects
    Call Class_Terminate
End Sub

Private Sub Class_Terminate()
    If Token = 0 Then Exit Sub
       GdiplusShutdown Token
       Token = 0
End Sub
