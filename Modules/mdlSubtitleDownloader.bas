Attribute VB_Name = "mdlSubtitleDownloader"
Option Explicit

Private Function ToHex(bytes() As Byte) As String
    
    Dim byteCurrByte As Byte
    
    For lngIter = 0 To UBound(bytes)
        byteCurrByte = bytes(lngIter)
        ToHex = ToHex & IIf(byteCurrByte < 10, "0", "") & Hex(bytes(lngIter))
        
    Next
    
End Function

Public Function ToHash(ByVal strFilePath As String) As Byte()
    
    Dim lngFileSize As Long
    Dim lngFileNum  As Long
    Dim bBuffer()   As Byte
    Dim lngIter     As Long
    Dim lngValue    As Currency
    Dim lngCurr     As Currency
    Dim lngCurPos   As Long
    lngValue = 0
    lngFileNum = FreeFile()
    Open strFilePath For Binary As #lngFileNum
    While ((lngIter < 8192))
        lngIter = lngIter + 1
        Get #lngFileNum, , lngCurr
        lngValue = CurrencyUnsignedAdd(lngValue, lngCurr)
    Wend
    Dim a As Variant
    lngCurPos = Seek(lngFileNum)
    Seek #lngFileNum, IIf(LOF(lngFileNum) - 65536 > 0, LOF(lngFileNum) - 65536, 0)
    
    lngIter = 0
    While ((lngIter < 8192))
        lngIter = lngIter + 1
        Get #lngFileNum, , lngCurr
        lngValue = CurrencyUnsignedAdd(lngValue, lngCurr)
    Wend
    
    Close #1
    
    ReDim bBuffer(8)
    CopyMem8 lngCurr, VarPtr(bBuffer(0))

    ToHash = bBuffer
End Function

Private Function CurrencyUnsignedAdd(src As Currency, target As Currency) As Currency
    Const CURR_MAX As Currency = 922337203685477.5807@
    If (CDbl(src) + CDbl(target) > CDbl(CURR_MAX)) Then
        CurrencyUnsignedAdd = src - CURR_MAX
        CurrencyUnsignedAdd = CurrencyUnsignedAdd + target
    Else
        CurrencyUnsignedAdd = src + target
    End If
End Function
