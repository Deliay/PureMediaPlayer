VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlphaPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lImage As Long

Private gImage As Long

Private gToken As Long

Private gStart As GdiplusStartupInput

Private phDC   As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long) As Long

Private Declare Function AlphaBlend _
                Lib "msimg32.dll" (ByVal hdcDest As Long, _
                                   ByVal nXOriginDest As Long, _
                                   ByVal nYOriginDest As Long, _
                                   ByVal nWidthDest As Long, _
                                   ByVal hHeightDest As Long, _
                                   ByVal hdcSrc As Long, _
                                   ByVal nXOriginSrc As Long, _
                                   ByVal nYOriginSrc As Long, _
                                   ByVal nWidthSrc As Long, _
                                   ByVal nHeightSrc As Long, _
                                   ByVal BLENDFUNCTION As Long) As Long

Private Type BLENDFUNCTION

    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte

End Type

Private Const AC_SRC_OVER = &H0

Public Event OnAlpha(lhDc As Long)

Public Event OnAlphaTab(lhDc As Long)

Public Event OnAlphaArea(lhDc As Long)

Public Event OverAlpha(bStatus As Boolean)

Public Event OverAlphaTab(bStatus As Boolean)

Public Event OverAlphaArea(bStatus As Boolean)

Public Event OnLoadFile(strPath As Long)

Public Event OverLoadFile(bStatus As Boolean)

Public Event OnDraw(lhDc As Long)

Public Event OverDraw(bStatus As Boolean)

Public Event OnLoadGDIPlus(siInput As GdiplusStartupInput)

Public Event OverLoadGDIPlus(bStatus As Boolean)

Public Event OnCreateImage()

Public Event OverCreateImage()

Public Event OnGetLin()

Public Event LineGot()

Public Event OverGetLin()

Public Event OnGetRow()

Public Event RowGot()

Public Event OverGotRow()

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Public Function AlphaArea(targetHDC As Long, _
                          value As Byte, _
                          sourceH As Long, _
                          sourceW As Long, _
                          Top As Long, _
                          Left As Long, _
                          selfH As Long, _
                          selfW As Long)
    
    Dim bf As BLENDFUNCTION, lBF As Long

    With bf
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = value
        .AlphaFormat = 0

    End With

    CopyMemory lBF, bf, 4
    AlphaBlend phDC, 0, 0, selfW \ 15, selfH \ 15, targetHDC, Left \ 15, Top \ 15, sourceW \ 15, sourceH \ 15, lBF

End Function

Public Function AlphaTab(targetHDC As Long, _
                         value As Byte, _
                         h As Long, _
                         w As Long, _
                         Top As Long, _
                         Left As Long)

    Dim bf As BLENDFUNCTION, lBF As Long

    With bf
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = value
        .AlphaFormat = 0

    End With
    
    CopyMemory lBF, bf, 4
    
    AlphaBlend phDC, 0, 0, w, h, targetHDC, Left \ 15, (Top \ 15), w, h, lBF

End Function

Public Function Alpha(targetHDC As Long, value As Byte, Top As Long, Left As Long)

    Dim bf As BLENDFUNCTION, lBF As Long

    With bf
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = value
        .AlphaFormat = 0

    End With

    CopyMemory lBF, bf, 4
    AlphaBlend phDC, 0, 0, ImageWidth, ImageHeight, targetHDC, IIf((Left \ 15) - 8 < 0, (Left \ 15), (Left \ 15) - 8), (Top \ 15), ImageWidth, ImageHeight, lBF

End Function

Public Property Get hDC() As Long
    hDC = phDC

End Property

Public Property Let hDC(lhDc As Long)

    If lhDc <> phDC Or phDC = 0 Then
        'If (And lHdc <> 0) Then
        GdiPlusExports.GdipDeleteGraphics gImage
        GdiPlusExports.GdipCreateFromHDC lhDc, gImage
        'End If
    ElseIf lhDc = 0 Then
        GdiPlusExports.GdipDeleteGraphics gImage
        gImage = 0

    End If

    phDC = lhDc

End Property

'
'Public Function hWnd2hDC2(lhDC As Long)
'
'    GdiPlusExports.GdipDeleteGraphics gImage
'    GdiPlusExports.GdipCreateBitmapFromHBITMAP lhDC, 0, gImage
'    'GdiPlusExports.GdipCreateFromHWND  lHdc, gImage
'    'GdiPlusExports.GdipCreateFromHDC lHdc, gImage
'    phDC = lhDC
'End Function

Public Function hWnd2hDC(lhDc As Long)

    If GetDC(lhDc) <> phDC Then
        If lhDc = 0 Then GoTo lhdce0
        If (phDC = 0 And GetDC(lhDc) <> 0) Then
            GdiPlusExports.GdipDeleteGraphics gImage
            'GdiPlusExports.GdipCreateBitmapFromHBITMAP lHdc, gImage
            GdiPlusExports.GdipCreateFromHDC GetDC(lhDc), gImage

        End If

    ElseIf lhDc = 0 Then
lhdce0:
        GdiPlusExports.GdipDeleteGraphics gImage
        gImage = 0

    End If

    phDC = GetDC(lhDc)

End Function

'
'Public Property Get ImageColor() As Long

'    Dim hTmpGp As Long
'    For k = 1 To 32
'        For i = 1 To 32
'            picColorPick.PSet (k, i), picColorRemix.Point(k, i)
'        Next
'    Next
'    hTmpGp = picColorPick.Point(1, 1)
'    Dim r As Long, g As Long, b As Long
'    r = hTmpGp Mod 256
'    g = (hTmpGp \ 256) Mod 256
'    b = hTmpGp \ 256 \ 256
'    hTmpGp = RGB(r * 0.9, g * 0.9, b * 0.9)
'    '
'End Property

'Public Property Get ImageColor() As Long
'
'End Property

Public Property Get Row(Y) As Long()

    Dim laRow() As Long

    Dim i       As Long

    If ImageHeight = 0 Then Exit Property
    ReDim laRow(0 To ImageHeight - 1)

    For i = 0 To UBound(laRow)
        laRow(i) = GetPixel(phDC, i + 1, Y)
    Next
    Row = laRow
    Erase laRow

End Property

Public Property Get lin(X) As Long()

    Dim laLin() As Long

    Dim i       As Long

    If ImageWidth = 0 Then Exit Property
    ReDim laLin(0 To ImageWidth - 1)

    For i = 0 To UBound(laLin)
        laLin(i) = GetPixel(phDC, X, i + 1)
    Next
    lin = laLin
    Erase laLin

End Property

Public Property Get Point(X, Y) As Long
    Point = GetPixel(phDC, X, Y)

End Property

Public Property Get ImageHeight() As Long

    If lImage <> 0 Then GdiPlusExports.GdipGetImageHeight lImage, ImageHeight

End Property

Public Property Get ImageWidth() As Long

    If lImage <> 0 Then GdiPlusExports.GdipGetImageWidth lImage, ImageWidth

End Property

Public Function LoadImage(Path)

    If phDC = 0 Then Exit Function
    If gImage = 0 Then GdiPlusExports.GdipCreateFromHDC phDC, gImage
    If lImage <> 0 Then GdiPlusExports.GdipDisposeImage lImage
    
    If GdiPlusExports.GdipLoadImageFromFile(CStr(Path), lImage) <> Ok Then

        'MsgBox "ERROR!"
    End If
    
    If GdipDrawImageRect(gImage, lImage, 0, 0, ImageWidth, ImageHeight) <> Ok Then

        'MsgBox "ERROR!"
    End If

End Function

Public Function LoadImageWH(Path, width As Long, height As Long)

    If phDC = 0 Then Exit Function
    If gImage = 0 Then GdiPlusExports.GdipCreateFromHDC phDC, gImage
    If lImage <> 0 Then GdiPlusExports.GdipDisposeImage lImage
    If GdiPlusExports.GdipLoadImageFromFile(CStr(Path), lImage) <> Ok Then

        'MsgBox "ERROR!"
    End If

    If GdipDrawImageRect(gImage, lImage, 0, 0, width, height) <> Ok Then

        'MsgBox "ERROR!"
    End If

End Function

Public Function Refresh()

    If (gImage = 0 Or lImage = 0) Then Exit Function
    If GdipDrawImageRect(gImage, lImage, 0, 0, width, height) <> Ok Then

        'MsgBox "ERROR!"
    End If

End Function

Public Function RefreshHW(width As Long, height As Long)

    If (gImage = 0 Or lImage = 0) Then Exit Function
    If GdipDrawImageRect(gImage, lImage, 0, 0, width, height) <> Ok Then
        MsgBox "err"

    End If

End Function

Private Sub Class_Initialize()
    gStart.GdiplusVersion = 1

    If GdiPlusExports.GdiplusStartup(gToken, gStart) <> Ok Then
        Err.Raise 5
        Unload Me

    End If
    
End Sub

Private Sub Class_Terminate()
    GdiPlusExports.GdipDeleteGraphics gImage
    GdiPlusExports.GdipDisposeImage lImage
    GdiPlusExports.GdiplusShutdown (gToken)
    
End Sub

