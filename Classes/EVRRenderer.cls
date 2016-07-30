VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EVRRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' IMFVideoDisplayControl
Const IID_IMFVideoDisplayCtrl = "{A490B1E4-AB84-4D31-A1B2-181E03B1077A}"

Const IID_IMFGetServices = "{FA993888-4383-415A-A930-DD472A8CF6F7}"

Const IID_MR_VIDEO_RENDER_SERVICE = "{1092a86c-ab1a-459a-a336-831fbc4d11ff}"

Private ifPlayback As IMFVideoDisplayControl

Private srcW       As Long, srcH As Long, srcL As Long, srcT As Long

Private ifServices As IMFGetService

Public Sub CreateInterface(baseFilter As olelib.IUnknown)

    Dim uuidGetServices As UUID, uuidVideoDisplayCtrl As UUID, uuidMVR As UUID

    Dim result          As Long, b As Variant
    
    olelib.CLSIDFromString IID_IMFGetServices, uuidGetServices
    olelib.CLSIDFromString IID_IMFVideoDisplayCtrl, uuidVideoDisplayCtrl
    olelib.CLSIDFromString IID_MR_VIDEO_RENDER_SERVICE, uuidMVR
    result = baseFilter.QueryInterface(uuidGetServices, ifServices)
        
    If (result <> 0) Then Err.Raise result, "Convert Interface"
     
    Set ifServices = baseFilter
    
    Set ifPlayback = ifServices.GetService(uuidMVR, uuidVideoDisplayCtrl)
    ifPlayback.SetAspectRatioMode 0
    ifPlayback.SetBorderColor 0

End Sub

Public Function GetInterface() As IMFVideoDisplayControl
    Set GetInterface = ifPlayback

End Function

Public Function SetPlayBackWindow(hWnd As Long)
    ifPlayback.SetVideoWindow hWnd

End Function

Public Function SetVideoSize(Left As Long, Top As Long, Width As Long, Height As Long)

    Dim rect       As EVRImport2.tagRECT

    Dim videoSizeH As Long, videoSizeW As Long

    rect.Left = Left
    rect.Top = Top
    
    rect.Right = Width + Left
    rect.Bottom = Top + Height
    GetVideoSize videoSizeW, videoSizeH

    If (Round((Width / Height), 1) = Round(videoSizeW / videoSizeH, 1)) Then
        ifPlayback.SetVideoPosition 0&, VarPtr(rect)

    End If

    If (srcW <> Width Or srcH <> Height Or srcT <> Top Or srcL <> Left) Then
        ifPlayback.RepaintVideo
        
        srcW = Width
        srcH = Height
        srcT = Top
        srcL = Left

    End If

End Function

Public Sub GetVideoSize(ByRef Width As Long, ByRef Heigth As Long)

    Dim rectSizeV As EVRImport2.tagSIZE, rectSizeVR As EVRImport2.tagSIZE

    ifPlayback.GetNativeVideoSize rectSizeV, rectSizeVR
    Width = rectSizeV.cx
    Heigth = rectSizeV.cy

End Sub
