#cs ----------------------------------------------------------------------------

 Title: GIFAnimatorPreview
 Version: 1.0.0
 AutoIt Version: 3.3.14.5
 Author:         Eduardo Mozart de Oliveira

 Script Function:
	File System Class for AutoIt.

#ce ----------------------------------------------------------------------------

#include <MsgBoxConstants.au3>

; _ArrayDisplay
#include <Array.au3>

#include <File.au3>

; _FolderExists
#include "AutoIt-FSOClass\FSOClass.au3"

; Yet Another Gif Example
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

Opt("MustDeclareVars", 1)

AutoItSetOption("MouseCoordMode", 2)
AutoItSetOption("TrayIconDebug", 1)
;OnAutoItExitRegister("OnAutoItExit")
HotKeySet("{ESCAPE}", "OnAutoItExit")

Global $GIFAnimatorPreviewDebug = True

; If _Singleton(@ScriptName, 1) = 0 Then
;   MsgBox($MB_SYSTEMMODAL, "Warning", "An occurrence of " & @ScriptName & " is already running")
;   Exit
; EndIf

Global $hWnd = Null
Global $GIFAnimatorPreviewhWnd = Null

Global $bPreview = False
Global $GIFOriginalPath = Null
Global $aGIF[2] ; Initiate the array.

GIFAnimatorPreviewParseArguments()

GIFAnimatorPreview($aGIF[1])

Func GIFAnimatorPreview($sGIF)
   ; https://www.autoitscript.com/forum/topic/125623-yet-another-gif-display/
   GUICreate("Yet Another Gif Example", 250, 400, -1, -1, BitOR($WS_MAXIMIZEBOX,$WS_SIZEBOX,$WS_THICKFRAME))
   ;create the object
   Local $oObj = ObjCreate("Shell.Explorer.2")
   Local $pwidth,$pheight
   _GetGifPixWidth_Height($sGIF, $pwidth, $pheight)
   Local $oObj_ctrl = GUICtrlCreateObj($oObj, 5, 5, $pwidth, $pheight)
   ;resize control when the window resizes
   GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
   ;restrict right click
   GUICtrlSetState(-1,$GUI_DISABLE)
   ;show the image
   ; https://www.autoitscript.com/forum/topic/85396-create-object-without-borders/
   ; https://stackoverflow.com/questions/20801411/background-image-height-width-doest-work-on-ie11
   Local $URL = "about:<html><body bgcolor='#efefef' scroll='no' style='border: 0; margin: 0 auto; text-align: center'><div style='position: fixed; top: -50%; left: -50%; width: 200%; height: 200%;'><img src='"&$sGIF&"' width='50%' height='50%' border='0' style='position: absolute; top: 0; left: 0; right: 0; bottom: 0; margin: auto;'></img></span></body></html>"
   ; Local $URL = "about:<html><body bgcolor='#efefef' scroll='no' style='border:0'><div style='position: fixed; top: -50%; left: -50%; width: 200%; height: 200%;'><img src='"&$sGIF&"' width='90%' height='100%' border='0'></img></span></body></html>"
   ; Local $URL = "about:<html><body bgcolor='#efefef' background='"&$sGIF&"' scroll='no' style='border:0;background-repeat:no-repeat;max-width:100px;max-height:100px'></body></html>"


   $oObj.Navigate($URL)
   GUISetState(@SW_SHOW)

   While 1
	  Sleep(10)
	  Switch GUIGetMsg()
		 Case -3
			$oObj=0 ;destroy object...
			Exit
	  EndSwitch
   WEnd

   ; GIFAnimatorPreviewOpen($sGIF)

   ; If Not WinActive(HWnd($GIFAnimatorPreviewhWnd)) Then WinActivate(HWnd($GIFAnimatorPreviewhWnd))

   ; _GUICtrlToolbar_ClickIndex($hToolbars, 16)
   ; WinWait("Preview")

   ; Local $aPreviewPos = ControlGetPos("Preview", "", "[CLASS:Static; INSTANCE:2]")
   ; If $GIFAnimatorPreviewDebug = True Then _ArrayDisplay($aPreviewPos, "aPreviewPos")

   ; GUICtrlSetResizing($phWnd, $GUI_DOCKAUTO)
   ; GUICtrlSetPos ($phWnd, $aPreviewPos[0], $aPreviewPos[1], 100, 100)
   ; If @error Then ConsoleWrite("Error on Resize.")
EndFunc    ;==>GIFAnimatorPreview

Func _FileSelectFolder()
    ; Create a constant variable in Local scope of the message to display in FileSelectFolder.
    Local Const $sMessage = "Select a folder"

    ; Display an open dialog to select a file.
    Local $sFileSelectFolder = FileSelectFolder($sMessage, "")
    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "No folder was selected. Aborting.")
		Exit 1
    Else
        ; Display the selected folder.
        ; MsgBox($MB_SYSTEMMODAL, "", "You choose the following folder:" & @CRLF & $sFileSelectFolder)
	 EndIf

	 Return $sFileSelectFolder
EndFunc   ;==>_FileSelectFolder

Func _WaitSelectDialog()
   While 1
	  ; Retrieve a list of window handles.
	  Local $aList = WinList()

	  ; Loop through the array displaying only visable windows with a title.
	  For $i = 1 To $aList[0][0]
		 If $aList[$i][0] <> "" And BitAND(WinGetState($aList[$i][1]), 2) Then
			; If $GIFAnimatorPreviewDebug = True Then ConsoleWrite("Title: " & $aList[$i][0] & @CRLF & "Handle: " & $aList[$i][1] & " Parent ID: " & _WinAPI_GetParent ( $aList[$i][1] ) & @CRLF)
			; If $GIFAnimatorPreviewDebug = True Then ConsoleWrite("If " & _WinAPI_GetParent ( $aList[$i][1] ) & " = " & $GIFAnimatorPreviewhWnd & " Then" & @CRLF)
			If _WinAPI_GetParent ( $aList[$i][1] ) = $GIFAnimatorPreviewhWnd Then
			   Return $aList[$i][1]
			EndIf
		 EndIf
	  Next
   WEnd
EndFunc   ;==>_WaitSelectDialog

Func GIFAnimatorPreviewParseArguments()
   Local $sFilePath = Null
   Local $sDirPath = Null

   ; https://www.autoitscript.com/forum/topic/794-parsing-command-line-args/
   Local $V_Arg =  "Preview *.GIF file(s)." & @CRLF & @CRLF
   $V_Arg = $V_Arg & "Valid Arguments are: " & @CRLF
   $V_Arg = $V_Arg & "    [Directory] - Search for *.GIF files into Directory." & @CRLF
   $V_Arg = $V_Arg & "    [File] - Open *.GIF file." & @CRLF
   $V_Arg = $V_Arg & "    /debug - Enable debug messages."
   ;$V_Arg = $V_Arg & "    /s       - Search for *.GIF Files Recursively into Subdirectories." & @CRLF
   ; retrieve commandline parameters
   For $x = 1 to $CmdLine[0]
	  Select
		 Case $CmdLine[$x] = "/debug"
			$GIFAnimatorPreviewDebug = True
		 ;Case $CmdLine[$x] = "/s"

		 Case $CmdLine[$x] = "/?" Or $CmdLine[$x] = "/h" Or $CmdLine[$x] = "/help"
			MsgBox( 1, "GIFAnimatorPreview", "" & $v_Arg,)
			Exit
		 Case Else
			If _FolderExists($CmdLine[$x]) Then
			   $sDirPath = $CmdLine[$x]
			   If $GIFAnimatorPreviewDebug = True Then ConsoleWrite("[Directory]: " & $sDirPath & @CRLF)
			ElseIf _FileExists($CmdLine[$x]) Then
			   $sFilePath = $CmdLine[$x]
			   If $GIFAnimatorPreviewDebug = True Then ConsoleWrite("[File]: " & $sFilePath & @CRLF)
			Else
			   MsgBox( 1, "GIFAnimatorPreview", "Wrong commandline argument: " & $CmdLine[$x] & @CRLF & $v_Arg,)
			   Exit
			EndIf
	  EndSelect
   Next


   If $sFilePath = Null And $sDirPath = Null Then
	  $sDirPath = _FileSelectFolder()
   EndIf

   If _FileExists($sFilePath) Then
	  $aGIF[0] = 1
	  $aGIF[1] = $sFilePath
   EndIf

   If _FolderExists($sDirPath) Then
	  ; $sTo = _FileSelectFolder()
	  ; https://www.autoitscript.com/forum/topic/142989-scan-for-certain-file-types/
	  ; Shows the filenames gif files in the current directory.
	  $aGIF = _FileListToArray($sDirPath, "*.gif*", 1)
	  If IsArray($aGif) Then
		 For $i = 1 to UBound($aGIF) -1
			$aGIF[$i] = $sDirPath & "\" & $aGIF[$i]
		 Next
	  EndIf
   EndIf

   If Not IsArray($aGIF) Then
	  MsgBox(0, "GIFAnimatorPreview", "No GIF files found at " & Quotes($sDirPath) & ". Aborting.")
	  Exit
   EndIf

   If $GIFAnimatorPreviewDebug = True Then
	  ; _ArrayDisplay($aGIF, "aGIF")
   EndIf
EndFunc   ;==>GIFAnimatorPreviewParseArguments

;===============================================================================
;
; Function Name:    _GetGifPixWidth_Height()
; Description:      return the size of a GIF image in pixels
; Parameter(s):     $s_gif      [required]      path and filename of the animated GIF
;
; Requirement(s):   #include <IE.au3>
; Return Value(s):
;                   $pwidth = width of the GIF in pixels
;                   $pheight = height of the GIF in pixels
; Author(s):        gafrost (https://www.autoitscript.com/forum/topic/33194-gifs-in-autoit/)
;
;===============================================================================
Func _GetGifPixWidth_Height($s_gif, ByRef $pwidth, ByRef $pheight)
    If FileGetSize($s_gif) > 9 Then
        Local $sizes = FileRead($s_gif, 10)
        ConsoleWrite("Gif version: " & StringMid($sizes, 1, 6) & @LF)
        $pwidth = Asc(StringMid($sizes, 8, 1)) * 256 + Asc(StringMid($sizes, 7, 1))
        $pheight = Asc(StringMid($sizes, 10, 1)) * 256 + Asc(StringMid($sizes, 9, 1))
        ConsoleWrite($pwidth & " x " & $pheight & @LF)
    EndIf
EndFunc   ;==>_GetGifPixWidth_Height

Func Quotes($sString)
   Return Chr(34) & $sString & Chr(34)
EndFunc   ;==>Quotes
