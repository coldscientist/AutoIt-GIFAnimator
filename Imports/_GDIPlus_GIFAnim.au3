;coded by UEZ build 2017-07-04
;requires 3.3.11.5+

#AutoIt3Wrapper_Version=b
#include-once
#include <GDIPlus.au3>
#include <Memory.au3>

;_GDIPlus_BitmapConvertTo8Bit
;_GDIPlus_GIFAnimCreateFile
;_GDIPlus_GIFAnimExtractAllFrames
;_GDIPlus_GIFAnimGetFrameCount
;_GDIPlus_GIFAnimGetFrameDelays
;_GDIPlus_GIFAnimGetFrameDelaysFromBinFile
;_GDIPlus_GIFAnimGetFrameDimensionsCount
;_GDIPlus_GIFAnimGetFrameDimensionsList
;_GDIPlus_GIFAnimSelectActiveFrame
;_GDIPlus_ImageGetColorPalette
;_GDIPlus_ImageGetColorPaletteSize
;_GDIPlus_ImageSaveAdd
;_GDIPlus_ImageSaveAddImage

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_GIFAnimGetFrameDimensionsCount
; Description ...: Gets the number of frame dimensions in this Image object.
; Syntax ........: _GDIPlus_GIFAnimGetFrameDimensionsCount($hImage)
; Parameters ....: $hImage              - A handle to an image / bitmap object
; Return values .: The number of frame dimensions in this Image object.
; Author ........: UEZ
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; ===============================================================================================================================
Func _GDIPlus_GIFAnimGetFrameDimensionsCount($hImage)
    Local Const $aResult = DllCall($__g_hGDIPDll, "int", "GdipImageGetFrameDimensionsCount", "handle", $hImage, "ulong*", 0)
    If @error Then Return SetError(@error, @extended, 0)
    If $aResult[0] Then Return SetError(10, $aResult[0], 0)
    Return $aResult[2]
EndFunc   ;==>_GDIPlus_GIFAnimGetFrameDimensionsCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_GIFAnimGetFrameDimensionsList
; Description ...: Gets the identifiers for the frame dimensions of this Image object which fills the GUID struct.
; Syntax ........: _GDIPlus_GIFAnimGetFrameDimensionsList($hImage, $iFramesCount)
; Parameters ....: $hImage              - A handle to an image / bitmap object
;                  $iFramesCount        - An integer value.
; Return values .: tagGUID struct
; Author ........: UEZ
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; ===============================================================================================================================
Func _GDIPlus_GIFAnimGetFrameDimensionsList($hImage, $iFramesCount)
    Local Const $tGUID = DllStructCreate($tagGUID)
    Local Const $aResult = DllCall($__g_hGDIPDll, "int", "GdipImageGetFrameDimensionsList", "handle", $hImage, "struct*", $tGUID, "uint", $iFramesCount)
    If @error Then Return SetError(@error, @extended, 0)
    If $aResult[0] Then Return SetError(10, $aResult[0], 0)
    Return $tGUID
EndFunc   ;==>_GDIPlus_GIFAnimGetFrameDimensionsList

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_GIFAnimGetFrameCount
; Description ...: Gets the frame count of the loaded gif by passing the GUID struct
; Syntax ........: _GDIPlus_GIFAnimGetFrameCount($hImage, $tGUID)
; Parameters ....: $hImage              - A handle to an image / bitmap object
;                  $tGUID               - A struct to a GUID that specifies the frame dimension.
; Return values .: The amount of frames from a GIF animated image handle
; Author ........: UEZ
; Modified ......:
; Remarks .......:
; Related .......: _GDIPlus_ImageLoadFromFile _GDIPlus_BitmapCreateFromFile
; Link ..........:
; ===============================================================================================================================
Func _GDIPlus_GIFAnimGetFrameCount($hImage, $tGUID)
    Local Const $aResult = DllCall($__g_hGDIPDll, "int", "GdipImageGetFrameCount", "handle", $hImage, "struct*", $tGUID, "ptr*", 0)
    If @error Then Return SetError(@error, @extended, 0)
    If $aResult[0] Then Return SetError(10, $aResult[0], 0)
    Return Int($aResult[3])
EndFunc   ;==>_GDIPlus_GIFAnimGetFrameCount

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_GIFAnimSelectActiveFrame
; Description ...: Selects the frame in this Image object specified by passing the GUID struct and current frame.
; Syntax ........: _GDIPlus_GIFAnimSelectActiveFrame($hImage, $tGUID, $iCurrentFrame)
; Parameters ....: $hImage              - A handle to an image / bitmap object
;                  $tGUID               - A struct to a GUID that specifies the frame dimension.
;                  $iCurrentFrame       - An integer value.
; Return values .: True or False on errors
; Author ........: UEZ
; Modified ......:
; Remarks .......:
; Related .......: _GDIPlus_ImageLoadFromFile _GDIPlus_BitmapCreateFromFile
; Link ..........:
; ===============================================================================================================================
Func _GDIPlus_GIFAnimSelectActiveFrame($hImage, $tGUID, $iCurrentFrame)
    Local Const $aResult = DllCall($__g_hGDIPDll, "int", "GdipImageSelectActiveFrame", "handle", $hImage, "struct*", $tGUID, "uint", $iCurrentFrame)
    If @error Then Return SetError(@error, @extended, 0)
    If $aResult[0] Then Return SetError(10, $aResult[0], 0)
    Return True
EndFunc   ;==>_GDIPlus_GIFAnimSelectActiveFrame

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_GIFAnimGetFrameDelays
; Description ...: Gets the delay of each frame from an image handle
; Syntax ........: _GDIPlus_GIFAnimGetFrameDelays($hImage, $iAnimFrameCount)
; Parameters ....: $hImage              - A handle to an image / bitmap object
;                  $iAnimFrameCount     - An integer value.
; Return values .: An array with the information about the delay of each frame or the error code
; Author ........: UEZ
; Modified ......:
; Remarks .......: If frame delays cannot be read try _GDIPlus_GIFAnimGetFrameDelaysFromBinFile instead
; Related .......: _GDIPlus_ImageLoadFromFile _GDIPlus_BitmapCreateFromFile _GDIPlus_ImageGetPropertyItem
; Link ..........:
; ===============================================================================================================================
Func _GDIPlus_GIFAnimGetFrameDelays($hImage, $iAnimFrameCount)
    If $iAnimFrameCount < 2 Then Return SetError(1, 0, 0)
    Local Const $GDIP_PROPERTYTAGFRAMEDELAY = 0x5100
    Local $tPropItem = __GDIPlus_ImageGetPropertyItem($hImage, $GDIP_PROPERTYTAGFRAMEDELAY)
    If IsDllStruct($tPropItem) And (Not @error) Then
        Local $iType = $tPropItem.type, $iLength, $tVal
        If $iType Then
            $iLength = $tPropItem.length
            Switch $iType
                Case 1
                    $tVal = DllStructCreate("byte delay[" & $iLength & "]", $tPropItem.value)
                Case 3
                    $tVal = DllStructCreate("short delay[" & Ceiling($iLength / 2) & "]", $tPropItem.value)
                Case 4
                    $tVal = DllStructCreate("long delay[" & Ceiling($iLength / 4) & "]", $tPropItem.value)
                Case Else
                    Return SetError(3, 0, 0)
            EndSwitch
            Local $aFrameDelays[Int($iAnimFrameCount)], $i
            For $i = 0 To UBound($aFrameDelays) - 1
                $aFrameDelays[$i] = $tVal.delay(($i + 1)) * 10
;~              ConsoleWrite($i & ": "& $aFrameDelays[$i] & @CRLF)
            Next
        EndIf
        Return $aFrameDelays
    EndIf
    Return SetError(2, 0, 0)
EndFunc   ;==>_GDIPlus_GIFAnimGetFrameDelays

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_GIFAnimGetFrameDelaysFromBinFile
; Description ...: Gets the delay of each frame from a binary gif file
; Syntax ........: _GDIPlus_GIFAnimGetFrameDelaysFromBinFile($binGIF, $iAnimFrameCount[, $iDelay = 10])
; Parameters ....: $binGIF              - A binary string with the GIF anim.
;                  $iAnimFrameCount     - An integer value.
; Return values .: An array with the information about the delay of each frame or the error code
; Author ........: UEZ
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; ===============================================================================================================================
Func _GDIPlus_GIFAnimGetFrameDelaysFromBinFile($binGIF, $iAnimFrameCount)
    If Not IsBinary($binGIF) Then Return SetError(1, 0, 0)
    If $iAnimFrameCount < 2 Then Return SetError(2, 0, 0)
    Local $aFrameDelays = StringRegExp($binGIF, "(?i)0021F904[[:xdigit:]]{2}([[:xdigit:]]{4})", 3)
    If @error Then Return SetError(3, 0, 0)
    Local Const $iDelay = 10
    For $i = 0 To UBound($aFrameDelays) - 1
        $aFrameDelays[$i] = $iDelay * Dec(StringRegExpReplace($aFrameDelays[$i], "([[:xdigit:]]{2})([[:xdigit:]]{2})", "$2$1"))
    Next
    If UBound($aFrameDelays) <> $iAnimFrameCount Then ReDim $aFrameDelays[$iAnimFrameCount]
    Return $aFrameDelays
EndFunc   ;==>_GDIPlus_GIFAnimGetFrameDelaysFromBinFile

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_GIFAnimExtractAllFrames
; Description ...: Extracts all frames from a GIF animation file with the option to resize the frames
; Syntax ........: _GDIPlus_GIFAnimExtractAllFrames($hImage, $sFilename[, $iJPGQual = 85[, $iW = 0[, $iH = 0[, $iResizeQual = 7[,
;                  $bReverse = False]]]]])
; Parameters ....: $hImage              - A handle to an image / bitmap object
;                  $sFilename           - A string value. Folders will be created if not existing.
;                  $iJPGQual            - [optional] An integer value. Default is 85.
;                  $iW                  - [optional] An integer value. Default is 0.
;                  $iH                  - [optional] An integer value. Default is 0.
;                  $iResizeQual         - [optional] An integer value. Default is 2.
;                  $bReverse            - [optional] A binary value. Default is False.
; Return values .: True or False on errors
; Author ........: UEZ
; Modified ......:
; Remarks .......: All frames will be extracted whereas the filename will be <filename>_XX.<ext>. XX is the amount of frames. If
;                  $bReverse is set True then the frames will be saved in reverse order. If $iW and $iH are zero then no resizing will be done.
; Related .......: _GDIPlus_EncodersGetCLSID _GDIPlus_ParamInit _GDIPlus_BitmapCreateFromScan0 _GDIPlus_GraphicsSetInterpolationMode _GDIPlus_ImageSaveToFile
; Link ..........:
; ===============================================================================================================================
Func _GDIPlus_GIFAnimExtractAllFrames($hImage, $sFilename, $iJPGQual = 85, $iW = 0, $iH = 0, $iResizeQual = 2, $bReverse = False)
    If $sFilename = "" Then Return SetError(1, @error, 0)
    Local Const $iAnimDimCount = _GDIPlus_GIFAnimGetFrameDimensionsCount($hImage)
    If @error Then Return SetError(2, @error, 0)
    Local Const $tGUID = _GDIPlus_GIFAnimGetFrameDimensionsList($hImage, $iAnimDimCount)
    If @error Then Return SetError(3, @error, 0)
    Local Const $iAnimFrameCount = _GDIPlus_GIFAnimGetFrameCount($hImage, $tGUID)
    If @error Then Return SetError(4, @error, 0)
    Local $sPath = StringRegExpReplace($sFilename, "(.+)\\.+", "$1")
    If StringLen($sPath) > 2 Then
        If Not FileExists($sPath) Then DirCreate($sPath)
    EndIf
    Local $sPrefixList = "jpg,png,bmp,gif,tif,", $sSuffix = StringRight($sFilename, 3)
    If Not StringInStr($sPrefixList, $sSuffix) Or StringMid($sFilename, StringLen($sFilename) - 3, 1) <> "." Then
        $sSuffix = "png"
        $sFilename &= ".png"
    EndIf
    Switch $sSuffix
        Case "jpg"
            Local $sCLSID = _GDIPlus_EncodersGetCLSID("JPG")
            Local $tParams = _GDIPlus_ParamInit(1)
            Local $tData = DllStructCreate("int Quality")
            Local $pData = DllStructGetPtr($tData)
            Local $pParams = DllStructGetPtr($tParams)
            $tData.Quality = $iJPGQual
            _GDIPlus_ParamAdd($tParams, $GDIP_EPGQUALITY, 1, $GDIP_EPTLONG, $pData)
    EndSwitch
    Local $sPrefix = StringTrimRight($sFilename, 4), $hFrame, $iCurrentFrame = 0, $i, $iRet, $hBitmap, $hGfx, $bError = False, $bResize = False
    If ($iW > 0) And ($iH > 0) And ($iW <> _GDIPlus_ImageGetWidth($hImage)) And ($iH <> _GDIPlus_ImageGetHeight($hImage)) Then
        $bResize = True
        $hBitmap = _GDIPlus_BitmapCreateFromScan0($iW, $iH)
        $hGfx = _GDIPlus_ImageGetGraphicsContext($hBitmap)
        _GDIPlus_GraphicsSetInterpolationMode($hGfx, $iResizeQual)
    EndIf
    Local $iFrame
    For $i = 0 To $iAnimFrameCount
        If $bReverse Then
            $iFrame = $iAnimFrameCount - $i
        Else
            $iFrame = $i
        EndIf
        _GDIPlus_GIFAnimSelectActiveFrame($hImage, $tGUID, $i)
        Switch $bResize
            Case False
                $hFrame = $hImage
            Case Else
                _GDIPlus_GraphicsClear($hGfx, 0x00000000)
                _GDIPlus_GraphicsDrawImageRect($hGfx, $hImage, 0, 0, $iW, $iH)
                $hFrame = $hBitmap
        EndSwitch
        Switch $sSuffix
            Case "jpg"
                $iRet = _GDIPlus_ImageSaveToFileEx($hFrame, $sPrefix & "_" & StringFormat("%0" & StringLen(Int($iAnimFrameCount)) & "i." & $sSuffix, $iFrame), $sCLSID, $pParams)
                If Not $iRet Then $bError = True
            Case Else
                $iRet = _GDIPlus_ImageSaveToFile($hFrame, $sPrefix & "_" & StringFormat("%0" & StringLen(Int($iAnimFrameCount)) & "i." & $sSuffix, $iFrame))
                If Not $iRet Then $bError = True
        EndSwitch
    Next
    If $bResize Then
        _GDIPlus_GraphicsDispose($hGfx)
        _GDIPlus_BitmapDispose($hBitmap)
    EndIf
    Return $bError
EndFunc   ;==>_GDIPlus_GIFAnimExtractAllFrames

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_GIFAnimCreateFile
; Description ...: Creates a GIF animation file
; Syntax ........: _GDIPlus_GIFAnimCreateFile($aImages, $sFilename[, $iDelay = 100])
; Parameters ....: $aImages             - An array of image handles (animation frames).
;                  $sFilename           - The filename of the GIF animation.
;                  $iDelay              - [optional] An integer value. Default is 100 ms per frame.
; Return values .: True or False on errors
; Author ........: UEZ
; Modified ......:
; Remarks .......: Vista or a higher operating system is required to create a GIF animation!
; Related .......: _GDIPlus_EncodersGetCLSID _GDIPlus_ParamInit _GDIPlus_ParamAdd _GDIPlus_ImageSaveToFileEx
; Link ..........:
; ===============================================================================================================================
Func _GDIPlus_GIFAnimCreateFile($aImages, $sFilename, $iDelay = 100)
    Local Const $GDIP_EVTFrameDimensionTime = 21
    Local $sCLSID = _GDIPlus_EncodersGetCLSID("GIF")
    Local $tMultiFrameParam = DllStructCreate("int;")

    DllStructSetData($tMultiFrameParam, 1, $GDIP_EVTMULTIFRAME)
    Local $tParams = _GDIPlus_ParamInit(1)
    _GDIPlus_ParamAdd($tParams, $GDIP_EPGSAVEFLAG, 1, $GDIP_EPTLONG, DllStructGetPtr($tMultiFrameParam))

    Local $hStream = _WinAPI_CreateStreamOnHGlobal()
    Local $tGUID = _WinAPI_GUIDFromString($sCLSID)
    _GDIPlus_ImageSaveToStream($aImages[1], $hStream, DllStructGetPtr($tGUID), DllStructGetPtr($tParams))

    DllStructSetData($tMultiFrameParam, 1, $GDIP_EVTFrameDimensionTime)
    $tParams = _GDIPlus_ParamInit(1)
    _GDIPlus_ParamAdd($tParams, $GDIP_EPGSAVEFLAG, 1, $GDIP_EPTLONG, DllStructGetPtr($tMultiFrameParam))

    Local $i
    For $i = 2 To $aImages[0] - 1
        _GDIPlus_ImageSaveAddImage($aImages[1], $aImages[$i], $tParams)
    Next

    DllStructSetData($tParams, 1, $GDIP_EVTLASTFRAME)
    $tParams = _GDIPlus_ParamInit(1)
    _GDIPlus_ParamAdd($tParams, $GDIP_EPGSAVEFLAG, 1, $GDIP_EPTLONG, DllStructGetPtr($tMultiFrameParam))
    _GDIPlus_ImageSaveAddImage($aImages[1], $aImages[$i], $tParams)

    DllStructSetData($tParams, 1, $GDIP_EVTFLUSH)
    $tParams = _GDIPlus_ParamInit(1)
    _GDIPlus_ParamAdd($tParams, $GDIP_EPGSAVEFLAG, 1, $GDIP_EPTLONG, DllStructGetPtr($tMultiFrameParam))
    _GDIPlus_ImageSaveAdd($aImages[$i], $tParams)

    Local $hMemory = _WinAPI_GetHGlobalFromStream($hStream)
    Local $iMemSize = _MemGlobalSize($hMemory)
    Local $pMem = _MemGlobalLock($hMemory)
    Local $tData = DllStructCreate("byte[" & $iMemSize & "]", $pMem)
    Local $bData = DllStructGetData($tData, 1)
    _WinAPI_ReleaseStream($hStream)
    _MemGlobalFree($hMemory)

    $bData = StringRegExpReplace($bData, "(?i)(0021F904[[:xdigit:]]{2})[[:xdigit:]]{4}", "${1}" & StringRegExpReplace(Hex(Int($iDelay / 10), 4), "([[:xdigit:]]{2})([[:xdigit:]]{2})", "$2$1"))
    Local $iExtended = @extended

    Local $hFile = FileOpen($sFilename, 2)
    If @error Then Return SetError(2, 0, False)
    FileWrite($hFile, Binary($bData))
    FileClose($hFile)

    If Not $iExtended Then Return SetError(1, 0, False)
    Return SetExtended($iExtended, True)
EndFunc   ;==>_GDIPlus_GIFAnimCreateFile

Func _GDIPlus_GIFAnimCreateFileFromImageFiles($aFrames, $sGIFFileName, $bReplay = True)
    Local $tagGIFHeader = "byte Header[6];byte Width[2];byte Height[2];byte PackedField[1];byte BackgroundColorIndex[1];byte PixelAspectRatio[1];"
    Local $tGIFHeader_1frame = DllStructCreate($tagGIFHeader & "byte ColorTable[768];")
    Local $hFile = _WinAPI_CreateFile($aFrames[0][0], 2, 2), $nBytes
    _WinAPI_ReadFile($hFile, DllStructGetPtr($tGIFHeader_1frame), DllStructGetSize($tGIFHeader_1frame), $nBytes)
    _WinAPI_CloseHandle($hFile)
    Local $iColorTableSize = 3 * 2 ^ (BitAND($tGIFHeader_1frame.PackedField, 7) + 1)
    Local $tGIFHeader_File = DllStructCreate($tagGIFHeader & "byte ColorTable[" & $iColorTableSize & "];byte ApplicationBlockExtension[18]")
    $tGIFHeader_File.Header = $tGIFHeader_1frame.Header
    $tGIFHeader_File.Width = $tGIFHeader_1frame.Width
    $tGIFHeader_File.Height = $tGIFHeader_1frame.Height
    $tGIFHeader_File.PackedField = $tGIFHeader_1frame.PackedField
    $tGIFHeader_File.BackgroundColorIndex = $tGIFHeader_1frame.BackgroundColorIndex
    $tGIFHeader_File.PixelAspectRatio = $tGIFHeader_1frame.PixelAspectRatio
    $tGIFHeader_File.ColorTable = BinaryMid($tGIFHeader_1frame.ColorTable, 1, $iColorTableSize)
    $tGIFHeader_File.ApplicationBlockExtension = Binary("0x21FF0B4E45545343415045322E3003010000")
    Local $bGIFHeader, $i, $b, $p, $d
    For $i = 1 To 8
        $bGIFHeader &= StringTrimLeft(DllStructGetData($tGIFHeader_File, $i), 2)
    Next

    For $i = 0 To UBound($aFrames) - 1
        $b = Binary(FileRead($aFrames[$i][0]))
        $p = Floor(StringInStr($b, "0021F904") / 2)
        $d = Hex(Dec($aFrames[$i][1] / 10), 4)
        If Not $p Then ContinueLoop
        $bGIFHeader &= StringMid(StringRegExpReplace(BinaryMid($b, $p, BinaryLen($b) - $p - 1), "(?i)0021F904([[:xdigit:]]{2})([[:xdigit:]]{4})(.*)", "0021F904${1}" & StringRight($d, 2) & StringLeft($d, 2) & "$3"), 3)
    Next

    $hFile = FileOpen($sGIFFileName, 18)
    FileWrite($hFile, Binary("0x" & $bGIFHeader & "3B"))
    FileClose($hFile)
EndFunc   ;==>_GDIPlus_GIFAnimCreateFileFromImageFiles

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_BitmapConvertTo8Bit
; Description ...: Converts a bitmap to a 8-bit image
; Syntax ........: _GDIPlus_BitmapConvertTo8Bit(Byref $hBitmap[, $iColorCount = 253[,$iDitherType = $GDIP_DitherTypeDualSpiral8x8[,
;                  $iPaletteType = $GDIP_PaletteTypeFixedHalftone252,[$bUseTransparentColor = True]]]])
; Parameters ....: $hBitmap              - A handle to an image / bitmap object
;                  $iColorCount          - [optional] An integer value. Default is 253.
;                  $iDitherType          - [optional] An integer value. Default is $GDIP_DitherTypeDualSpiral8x8. -> http://msdn.microsoft.com/en-us/library/ms534106(v=vs.85).aspx
;                  $iPaletteType         - [optional] An integer value. Default is $GDIP_PaletteTypeFixedHalftone252 . -> http://msdn.microsoft.com/en-us/library/ms534159(v=vs.85).aspx
;                  $bUseTransparentColor - [optional] A binary value. Default is True.
; Return values .: True or False on errors
; Author ........: UEZ
; Modified ......:
; Remarks .......: Vista or a higher operating system is required
; Related .......: _GDIPlus_PaletteInitialize _GDIPlus_BitmapConvertFormat _GDIPlus_ImageLoadFromFile _GDIPlus_BitmapCreateFromScan0
; Link ..........: http://msdn.microsoft.com/en-us/library/windows/desktop/ms534106(v=vs.85).aspx)
; ===============================================================================================================================
Func _GDIPlus_BitmapConvertTo8Bit(ByRef $hBitmap, $iColorCount = 256, $iDitherType = $GDIP_DitherTypeDualSpiral8x8, $iPaletteType = $GDIP_PaletteTypeFixedHalftone252, $bUseTransparentColor = True)
    $iColorCount = ($iColorCount > 2 ^ 8) ? 2 ^ 8 : $iColorCount
    Local $tPalette = _GDIPlus_PaletteInitialize(256, $iPaletteType, $iColorCount, $bUseTransparentColor, $hBitmap)
    If @error Then Return SetError(1, @error, 0)
    Local $iRet = _GDIPlus_BitmapConvertFormat($hBitmap, $GDIP_PXF08INDEXED, $iDitherType, $iPaletteType, $tPalette)
    If @error Then Return SetError(2, @error, 0)
    Return $iRet
EndFunc   ;==>_GDIPlus_BitmapConvertTo8Bit

Func _GDIPlus_BitmapConvertToXBit(ByRef $hBitmap, $iColorCount = 16, $iPixelFormat = $GDIP_PXF04INDEXED, $iDitherType = $GDIP_DitherTypeDualSpiral8x8, $iPaletteType = $GDIP_PaletteTypeFixedHalftone252, $bUseTransparentColor = False)
    Switch $iPixelFormat
        Case $GDIP_PXF08INDEXED
            $iColorCount = ($iColorCount > 2 ^ 8) ? 2 ^ 8 : $iColorCount
        Case $GDIP_PXF04INDEXED
            $iColorCount = ($iColorCount > 2 ^ 4) ? 2 ^ 4 : $iColorCount
        Case $GDIP_PXF01INDEXED
            $iPaletteType = $GDIP_PaletteTypeFixedBW
            $iColorCount = 2
        Case Else
            $iPixelFormat = $GDIP_PXF04INDEXED
            $iColorCount = 16
            $iDitherType = $GDIP_DitherTypeDualSpiral8x8
            $iPaletteType = $GDIP_PaletteTypeFixedHalftone252
    EndSwitch
    Local $tPalette = _GDIPlus_PaletteInitialize(256, $iPaletteType, $iColorCount, $bUseTransparentColor, $hBitmap)
    If @error Then Return SetError(1, @error, 0)
    Local $iRet = _GDIPlus_BitmapConvertFormat($hBitmap, $iPixelFormat, $iDitherType, $iPaletteType, $tPalette)
    If @error Then Return SetError(2, @error, 0)
    Return $iRet
EndFunc   ;==>_GDIPlus_BitmapConvertToXBit

Func _GDIPlus_ImageGetColorPalette($hImage, ByRef $tColorPalette, $iSize)
    Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetImagePalette", "handle", $hImage, "struct*", $tColorPalette, "uint", $iSize)
    If @error Then Return SetError(@error, @extended, 0)
    If $aResult[0] Then Return SetError(10, $aResult[0], 0)
    Return True
EndFunc   ;==>_GDIPlus_ImageGetColorPalette

Func _GDIPlus_ImageGetColorPaletteSize($hImage)
    Local $aResult = DllCall($__g_hGDIPDll, "uint", "GdipGetImagePaletteSize", "handle", $hImage, "uint*", 0)
    If @error Then Return SetError(@error, @extended, 0)
    If $aResult[0] Then Return SetError(10, $aResult[0], 0)
    Return $aResult[2]
EndFunc   ;==>_GDIPlus_ImageGetColorPaletteSize

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_ImageGetPropertyItem
; Description ...: Gets a specified property item (piece of metadata) from this Image object.
; Syntax ........: __GDIPlus_ImageGetPropertyItem($hImage, $iPropID)
; Parameters ....: $hImage              - A handle to an image object.
;                  $iPropID             - An integer that identifies the property item to be retrieved.
; Return values .: $tagGDIPPROPERTYITEM structure or 0 on errors
; Author ........: UEZ
; Modified ......:
; Remarks .......:
; Related .......: _GDIPlus_ImageLoadFromFile _GDIPlus_ImageLoadFromStream
; Link ..........: Property Item Descriptions -> http://msdn.microsoft.com/en-us/library/windows/desktop/ms534416(v=vs.85).aspx
; ===============================================================================================================================
Func __GDIPlus_ImageGetPropertyItem($hImage, $iPropID)
    Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetPropertyItemSize", "handle", $hImage, "uint", $iPropID, "ulong*", 0)
    If @error Then Return SetError(@error, @extended, 0)
    If $aResult[0] Then Return SetError(10, $aResult[0], 0)
    Local Static $tBuffer ;why static? because otherwise it would crash when running it as x64 exe (workaround)
    $tBuffer = DllStructCreate("byte[" & $aResult[3] & "]")
    $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetPropertyItem", "handle", $hImage, "uint", $iPropID, "ulong", $aResult[3], "struct*", $tBuffer)
    If @error Then Return SetError(@error, @extended, 0)
    If $aResult[0] Then Return SetError(11, $aResult[0], 0)
    Local Const $tagGDIPPROPERTYITEM = "uint id;ulong length;word type;ptr value"
    Local $tPropertyItem = DllStructCreate($tagGDIPPROPERTYITEM, DllStructGetPtr($tBuffer))
    If @error Then Return SetError(20, $aResult[0], 0)
    Return $tPropertyItem
EndFunc   ;==>__GDIPlus_ImageGetPropertyItem

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_ImageSaveAddImage
; Description ...: Adds a frame to a file or stream specified in a previous call to the _GDIP_SaveImageToFile or _GDIP_SaveImageToStream functions.
; Syntax ........: _GDIPlus_ImageSaveAddImage($hImage, $hImageFrame, $tParams)
; Parameters ....: $hImage              - A handle to an image object.
;                  $hImageFrame         - A handle to an image object that holds the frame to be added.
;                  $tParams             - A dll struct to an EncoderParameters structure that holds parameters required by the image encoder
;                                         used by the save-add operation.
; Return values .: True or False on errors
; Author ........: UEZ
; Modified ......:
; Remarks .......:
; Related .......: _GDIPlus_ImageSaveAdd _GDIPlus_ParamInit _GDIP_SaveImageToFile _GDIP_SaveImageToStream _GDIPlus_ImageLoadFromFile _GDIPlus_BitmapCreateFromScan0
; ===============================================================================================================================
;~ Func _GDIPlus_ImageSaveAddImage($hImage, $hImageFrame, $tParams)
;~  Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipSaveAddImage", "handle", $hImage, "handle", $hImageFrame, "struct*", $tParams)
;~  If @error Then Return SetError(@error, @extended, False)
;~  If $aResult[0] Then Return SetError(10, $aResult[0], False)
;~  Return True
;~ EndFunc   ;==>_GDIPlus_ImageSaveAddImage

; #FUNCTION# ====================================================================================================================
; Name ..........: _GDIPlus_ImageSaveAdd
; Description ...: Adds a frame to a file or stream specified in a previous call to the _GDIP_SaveImageToFile or _GDIP_SaveImageToStream functions.
;                  Use this method to save selected frames from a multiple-frame image to another multiple-frame image.
; Syntax ........: _GDIPlus_ImageSaveAdd($hImage, $tParams)
; Parameters ....: $hImage              - A handle to an image object.
;                  $tParams             - A dll struct to a encoder parameter list structure ($tagGDIPENCODERPARAMS).
; Return values .: True or False on errors
; Author ........: UEZ
; Modified ......:
; Remarks .......:
; Related .......: _GDIP_SaveImageToFile _GDIP_SaveImageToStream _GDIPlus_ParamInit _GDIPlus_ImageLoadFromFile _GDIPlus_BitmapCreateFromScan0
; ===============================================================================================================================
;~ Func _GDIPlus_ImageSaveAdd($hImage, $tParams)
;~  Local $aResult
;~  $aResult = DllCall($__g_hGDIPDll, "int", "GdipSaveAdd", "handle", $hImage, "struct*", $tParams)
;~  If @error Then Return SetError(@error, @extended, False)
;~  If $aResult[0] Then Return SetError(10, $aResult[0], False)
;~  Return True
;~ EndFunc   ;==>_GDIPlus_ImageSaveAdd