#cs ----------------------------------------------------------------------------

 Title: GIFAnimator
 Version: 1.0.1.0
 AutoIt Version: 3.3.14.5
 Author:         Eduardo Mozart de Oliveira

 Script Function:
	File System Class for AutoIt.

#ce ----------------------------------------------------------------------------

#include <MsgBoxConstants.au3>

#include <WinAPIProc.au3>

#include <Array.au3>
#include <File.au3>

; _GUICtrlToolbar
#include <GuiToolbar.au3>
; GetAllWindowsControls
; #include "Functions.au3"

#include <WinAPISysWin.au3>

; _Singleton
#include <Misc.au3>

#include <GuiTab.au3>

; _WinAPI_IsHungAppWindow
#include <WinAPISysWin.au3>

; _GDIPlus
#include "Imports\_GDIPlus_GIFAnim.au3"

; _GetHwndFromPID, _WinChildDialogs
#include "Imports\AutoIt-ProcessClass\ProcessClass.au3"

Opt("MustDeclareVars", 1)

AutoItSetOption("MouseCoordMode", 2)
AutoItSetOption("TrayIconDebug", 1)
;OnAutoItExitRegister("OnAutoItExit")
HotKeySet("^e", "GIFAnimatorQuit")

Global $GIFAnimatorDebug = True

Global $GIFAnimatorPath = Null

If _Singleton(@ScriptName, 1) = 0 Then
   MsgBox($MB_ICONERROR, "GIFAnimator", "An occurrence of " & @ScriptName & " is already running. Aborting.")
   Exit
EndIf

If ProcessExists("GIFAnimator.exe") Then
   ; WinActivate("FileSaver")
   ; If WinExists("[CLASS:Static; INSTANCE:3]") Then Send("{ENTER}")

   ; WinKill("Microsoft Gif Animator")

   ; https://www.autoitscript.com/forum/topic/188748-how-to-get-target-path-of-a-process/
   Local $iID, $file, $parentID, $parentFile, $a_process = ProcessList("GIFAnimator.exe")
   For $i = 1 To $a_process[0][0]
	   $iID = $a_process[$i][1]
   ;   $parentID = _WinAPI_GetParentProcess($iID)
	   $file = _WinAPI_GetProcessFileName($iID)
	   $GIFAnimatorPath = $file
	   ProcessClose($iID)
   ;   FileSetAttrib($file, "-RASHNOT")
   ;   FileDelete($file)
   ;   $parentFile = _WinAPI_GetProcessFileName($parentID)
   ;~  ProcessClose($parentID)
   ;~  FileSetAttrib($parentFile, "-RASHNOT")
   ;~  FileDelete($parentFile)
   ;   MsgBox(0, $parentID, $parentFile)
	Next
EndIf

If $GIFAnimatorPath = Null Then
   Local $ImageComposerAddIn = "HKLM\SOFTWARE\Wow6432Node\Microsoft\Microsoft Image Composer\AddIn"
   $GIFAnimatorPath = RegRead($ImageComposerAddIn, "Microsoft GIF Animator")
   If @error Then
	  ; Can't open Key
	  ; ConsoleWrite(@error)

	  Global $ProgramFilesDir = EnvGet('ProgramFiles(x86)')
	  If Not $ProgramFilesDir Then $ProgramFilesDir = @ProgramFilesDir
	  If FileExists($ProgramFilesDir & "\Microsoft GIF Animator\GIFAnimator.exe") Then
		 $GIFAnimatorPath = $ProgramFilesDir & "\Microsoft GIF Animator\GIFAnimator.exe"
	  Else
		 MsgBox(BitOr($MB_ICONERROR, $MB_SYSTEMMODAL), "GIFAnimator.au3", "GIFAnimator.exe could not be found at " & _Quotes($ProgramFilesDir & "\Microsoft GIF Animator\GIFAnimator.exe") & " or " & $ImageComposerAddIn & ". Aborting.")
		 Exit
	  EndIf
   EndIf
EndIf

Global $hWnd = Null
Global $GIFAnimatorhWnd = Null

Global $GIFOriginalPath = Null
Global $aGIF[2] ; Initiate the array.

GIFAnimatorParseArguments()

GIFAnimator()

;===============================================================================
;
; Function Name:    GIFAnimator()
; Description:      Automatize Microsoft GIF Animator UI interface
;
; Requirement(s):   #include <GuiToolbar.au3>
;
; Return Value(s):
;
; Author(s):        Eduardo Mozart de Oliveira
;
;===============================================================================
Func GIfAnimator()
   Local $PID = Run($GIFAnimatorPath)
   $hWnd = WinWait("Microsoft Gif Animator")
   $GIFAnimatorhWnd = String($hWnd)
   AdlibRegister("GIFAnimatorClose")
   If $GIFAnimatorDebug = True Then ConsoleWrite("GIFAnimator.exe hWnd: " & $GIFAnimatorhWnd & @CRLF)
   Global $hToolbar = ControlGetHandle(HWnd($GIFAnimatorhWnd), "", "[CLASS:ToolbarWindow32; INSTANCE:1]")

   ; https://www.autoitscript.com/forum/topic/136041-solved-_filelisttoarray-need-an-explanation/
   For $i = 1 to UBound($aGIF) -1
	  ConsoleWrite( "(" & ($i) & " of " & $aGIF[0] & ") " & $aGIF[$i] & @crlf)
	  Local $aDuration = _GetGIFDuration($aGIF[$i])
	  ; _ArrayDisplay($aDuration)
	  If Not IsArray($aDuration) Then ContinueLoop
	  Local $iDuration = $aDuration[0] / 10
	  ConsoleWrite("Duration: " & $iDuration & @CRLF)
	  If $iDuration <= 5 Then
	  	 ContinueLoop ; Do Nothing
	  EndIf

	  ; https://www.autoitscript.com/forum/topic/120516-test-for-window-responsiveness/
	  While 1
		 If $GIFAnimatorDebug Then ConsoleWrite("_WinAPI_IsHungAppWindow: " & _WinAPI_IsHungAppWindow(HWnd($GIFAnimatorhWnd)) & @CRLF)
		 If _WinAPI_IsHungAppWindow(HWnd($GIFAnimatorhWnd)) = False Then
			ExitLoop
		 EndIf
		 Sleep(5000)
	  WEnd

	  GIFAnimatorOpen($aGIF[$i])

	  ; Animation Width / Height
	  ; GIFAnimatorWidthHeight()

	  ; Duration
	  ControlCommand(HWnd($GIFAnimatorhWnd), "", "[CLASS:SysTabControl32; INSTANCE:1]", "TabRight", "")
	  ControlCommand(HWnd($GIFAnimatorhWnd), "", "[CLASS:SysTabControl32; INSTANCE:1]", "TabRight", "")

	  Local $bhWnd = Null
	  While 1
		 $bhWnd = ControlGetHandle(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:3]")
		 If $GIFAnimatorDebug = True Then ConsoleWrite("Button3 Handle: " & String($bhWnd) & @CRLF)
		 Sleep(100)
		 ; https://stackoverflow.com/questions/14571668/autoit-wait-for-a-control-element-to-appear
		 If $bhWnd Then
			; we got the handle, so the button is there
			; now do whatever you need to do
			ExitLoop
		 EndIf
	  WEnd
	  ControlFocus(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:3]")

	  ; Local $iDuration = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:3]") ; 1 = 10 msec
	  ; If @error Then
	  ; 	ConsoleWrite("Error reading Duration data.")
	  ; 	Exit
	  ; Else
	  ;		ConsoleWrite("Duration: " & $iDuration & @CRLF)
	  ; EndIf

	  ; https://www.autoitscript.com/autoit3/docs/intro/lang_operators.htm
	  ; If $iDuration <= 5 Then
	  ;	 ContinueLoop ; Do Nothing
	  ; EndIf

	  ; Select All (CTRL + L)
	  ; Insert isn't enable when Select All is clicked
	  Local $bToolbarInsert = Null
	  While 1
		 If Not WinActive(HWnd($GIFAnimatorhWnd)) Then WinActivate(HWnd($GIFAnimatorhWnd))
		 ; _ArrayDisplay(GetAllWindowsControls(HWnd($GIFAnimatorhWnd)))
		 ; $bToolbarInsert = ControlCommand(HWnd($GIFAnimatorhWnd), "", $hToolbar, "IsEnabled", _GUICtrlToolbar_IndexToCommand($hToolbar, 3))
		 Local $aToolbarInsert = _GUICtrlToolbar_GetButtonInfo($hToolbar, _GUICtrlToolbar_IndexToCommand($hToolbar, 3)) ; _GUICtrlToolbar_GetButtonState
		 ; _ArrayDisplay($aToolbarInsert)
		 Local $iToolbarInsertState = $aToolbarInsert[1]
		 If $GIFAnimatorDebug Then ConsoleWrite("Toolbar Button 'Insert' (Command ID: " & _GUICtrlToolbar_IndexToCommand($hToolbar, 3) & ") State: " & $iToolbarInsertState & @CRLF)
		 If $iToolbarInsertState = 0 Then ExitLoop ; Insert Button Disabled
		 _GUICtrlToolbar_ClickIndex($hToolbar, 11)
		 ; ControlCommand(HWnd($GIFAnimatorhWnd), "", $hToolbar, "SendCommandID", _GUICtrlToolbar_IndexToCommand($hToolbar, 11))
		 Sleep(100)
	  WEnd

	  ; Change Duration
	  If Not WinActive(HWnd($GIFAnimatorhWnd)) Then WinActivate(HWnd($GIFAnimatorhWnd))
	  If $GIFAnimatorDebug = True Then ConsoleWrite("New Duration: " & Ceiling($iDuration / 2) & @CRLF)
	  ControlSetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:3]", Ceiling($iDuration / 2))

	  ; Create backup file
	  Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	  Local $aPathSplit = _PathSplit($aGIF[$i], $sDrive, $sDir, $sFileName, $sExtension)
	  ; _ArrayDisplay($aPathSplit, "_PathSplit of " & @ScriptFullPath)

	  Local $sBackupFile = $GIFOriginalPath & "\" & $sFileName & $sExtension
	  If FileExists($sBackupFile) Then
		 For $j = 0 to 4294967295
			$sBackupFile = $GIFOriginalPath & "\" & $sFileName & " (" & $j & ")" & $sExtension
			If Not FileExists($sBackupFile) Then
			   ExitLoop
			EndIf
		 Next
	  EndIf

	  FileCopy($aGIF[$i], $sBackupFile)
	  If Not FileExists($sBackupFile) Then
		 MsgBox(0, "GIFAnimator", "Could not create backup file in " & _Quotes($sBackupFile) & ". Aborting.")
		 Exit
	  Else
		 FileDelete($aGIF[$i])
		 If FileExists($aGIF[$i]) Then
			MsgBox(0, "GIFAnimator", "Could not delete " & _Quotes($aGIF[$i]) & " file. Aborting.")
			Exit
		 EndIf
	  EndIf

	  ; Save As (Ctrl + A)
	  GIFAnimatorSaveAsDialog()

	  ; Writing
	  ; http://www.autoitscript.com/forum/topic/95905-wait-until-button-text-ok/
	  Local $sReady = Null
	  While 1
		 ; https://www.autoitscript.com/forum/topic/4646-text-with-edit-and-static-classes/
		 $sReady = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Static; INSTANCE:25]")
		 If $GIFAnimatorDebug = True Then ConsoleWrite("Ready (Writing): " & $sReady & @CRLF)
		 Sleep(100)
		 If $sReady = "Ready" Then
			If Not WinActive(HWnd($GIFAnimatorhWnd)) Then WinActivate(HWnd($GIFAnimatorhWnd))
			ExitLoop
		 EndIf
	  WEnd
   Next
EndFunc    ;==>GIFAnimator

Func GIFAnimatorSaveAsDialog()
   If Not WinActive(HWnd($GIFAnimatorhWnd)) Then WinActivate(HWnd($GIFAnimatorhWnd))
   _GUICtrlToolbar_ClickIndex($hToolbar, 4)
   Local $ohWnd = _WaitChildDialog($GIFAnimatorhWnd)
   If @error Then
	  GIFAnimatorSaveAsDialog()
	  Return
   EndIf
   WinActivate($ohWnd)
   ; ControlSend{"{ENTER}")
   ControlClick($ohWnd, "", "Button2")
EndFunc

Func GIFAnimatorOpen($sFilePath)
   If Not WinActive(HWnd($GIFAnimatorhWnd)) Then WinActivate(HWnd($GIFAnimatorhWnd))

   ; New (CTRL + N)
   ; https://www.autoitscript.com/forum/topic/155068-toolsbar-understanding-how-to-access-these-controls/
   _GUICtrlToolbar_ClickIndex($hToolbar, 0)

   ; Open (CTRL + O)
   _GUICtrlToolbar_ClickIndex($hToolbar, 1)

   Local $ohWnd = _WaitChildDialog($GIFAnimatorhWnd)
   If @error Then
	  If Not WinExists($ohWnd) Then ; Valid HWnd
		 GIFAnimatorOpen($sFilePath)
	  EndIf
	  Return
   EndIf
   ; Probably user intervation, so let's ignore it
   If WinGetTitle($ohWnd) = "File Not Saved" Then
	  ControlClick($ohWnd, "", "Button2")
   EndIf

   ; https://www.autoitscript.com/forum/topic/166908-cant-controll-class32770-windows-7/
   If Not WinActive($ohWnd) Then WinActivate($ohWnd)
   WinWaitActive($ohWnd)
   ControlSetText($ohWnd, "", "[CLASS:Edit; INSTANCE:1]", $sFilePath) ; Set the edit control in Open with some text. The handle returned by WinWait is used for the "title" parameter of ControlSetText.
   ControlClick($ohWnd, "", "Button2")

   ; Reading
   ; http://www.autoitscript.com/forum/topic/95905-wait-until-button-text-ok/
   Local $sReady = Null
   While 1
	  ; https://www.autoitscript.com/forum/topic/4646-text-with-edit-and-static-classes/
	  $sReady = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Static; INSTANCE:8]")
	  If _IsEmpty($sReady) Then $sReady = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Static; INSTANCE:25]")
	  If $GIFAnimatorDebug = True Then ConsoleWrite("Ready (Open): " & $sReady & @CRLF)
	  Sleep(100)
	  If $sReady = "Ready" Then
		 $sReady = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Static; INSTANCE:25]")
		 ExitLoop
	  EndIf
   WEnd

   If Not StringInStr(WinGetTitle(HWnd($GIFAnimatorhWnd)), ".gif") Then
	  GIFAnimatorOpen($sFilePath)
	  Return
   EndIf

   Local $iTabCount = 1
   While 1
	  $iTabCount = _GUICtrlTab_GetItemCount(ControlGetHandle(HWnd($GIFAnimatorhWnd), "", "[CLASS:SysTabControl32; INSTANCE:1]"))
	  If $GIFAnimatorDebug Then ConsoleWrite("Tab Count: " & $iTabCount & @CRLF)
	  Sleep(100)
	  If $iTabCount >= 2 Then
		 ExitLoop
	  EndIf
   WEnd
EndFunc   ;==>GIFAnimatorOpen

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

Func GIFAnimatorParseArguments()
   Local $sFilePath = Null
   Local $sDirPath = Null

   ; https://www.autoitscript.com/forum/topic/794-parsing-command-line-args/
   Local $V_Arg = "Valid Arguments are: " & @CRLF
   $V_Arg = $V_Arg & "    [Directory] - Search for *.GIF files into Directory." & @CRLF
   $V_Arg = $V_Arg & "    [File] - Open *.GIF file." & @CRLF
   $V_Arg = $V_Arg & "    /debug - Enable debug messages."
   ;$V_Arg = $V_Arg & "    /s       - Search for *.GIF Files Recursively into Subdirectories." & @CRLF
   ; retrieve commandline parameters
   For $x = 1 to $CmdLine[0]
	  Select
		 Case $CmdLine[$x] = "/debug"
			$GIFAnimatorDebug = True
		 ;Case $CmdLine[$x] = "/s"

		 Case $CmdLine[$x] = "/?" Or $CmdLine[$x] = "/h" Or $CmdLine[$x] = "/help"
			MsgBox( 1, "GIFAnimator", "" & $v_Arg)
			Exit
		 Case Else
			If _FolderExists($CmdLine[$x]) Then
			   $sDirPath = $CmdLine[$x]
			   If $GIFAnimatorDebug = True Then ConsoleWrite("[Directory]: " & $sDirPath & @CRLF)
			ElseIf FileExists($CmdLine[$x]) Then
			   $sFilePath = $CmdLine[$x]
			   If $GIFAnimatorDebug = True Then ConsoleWrite("[File]: " & $sFilePath & @CRLF)
			Else
			   MsgBox( 1, "GIFAnimator", "Wrong commandline argument: " & $CmdLine[$x] & @CRLF & $v_Arg)
			   Exit
			EndIf
	  EndSelect
   Next


   If $sFilePath = Null And $sDirPath = Null Then
	  $sDirPath = _FileSelectFolder()
   EndIf

   If FileExists($sFilePath) Then
	  $aGIF[0] = 1
	  $aGIF[1] = $sFilePath
   EndIf

   If _FolderExists($sDirPath) Then
	  ; $sTo = _FileSelectFolder()
	  ; https://www.autoitscript.com/forum/topic/142989-scan-for-certain-file-types/
	  ; Shows the filenames gif files in the current directory.
	  $aGIF = _FileListToArray($sDirPath, "*.gif*", 1)
	  If IsArray($aGIF) Then
		 For $i = 1 to UBound($aGIF) -1
			$aGIF[$i] = $sDirPath & "\" & $aGIF[$i]
		 Next

		 $GIFOriginalPath = $sDirPath & "\GIF (Original)"
		 If Not _FolderExists($GIFOriginalPath) Then
			If Not DirCreate($GIFOriginalPath) Then
			   MsgBox(0, "GIFAnimator", "Could not create " & _Quotes($GIFOriginalPath) & " directory. Aborting.")
			   Exit
			EndIf
		 EndIf
	  EndIf
   EndIf

   If Not IsArray($aGIF) Then
	  MsgBox(0, "GIFAnimator", "No GIF files found at " & _Quotes($sDirPath) & ". Aborting.")
	  Exit
   EndIf

   If $GIFAnimatorDebug = True Then
	  ; _ArrayDisplay($aGIF, "aGIF")
   EndIf
EndFunc   ;==>GIFAnimatorParseArguments

Func GIFAnimatiorWidthHeight()
   ; Image Width / Height
   ControlCommand(HWnd($GIFAnimatorhWnd), "", "[CLASS:SysTabControl32; INSTANCE:1]", "TabRight", "")

   Local $whWnd = Null
   Local $hhWnd = Null
   While 1
	  $whWnd = ControlGetHandle(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:1]")
	  $hhWnd = ControlGetHandle(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:2]")
	  If $GIFAnimatorDebug = True Then ConsoleWrite("Width Handle: " & String($whWnd) & @CRLF)
	  If $GIFAnimatorDebug = True Then ConsoleWrite("Height Handle: " & String($hhWnd) & @CRLF)
	  Sleep(100)
	  ; https://stackoverflow.com/questions/14571668/autoit-wait-for-a-control-element-to-appear
	  If $whWnd And $hhWnd Then
		 ; we got the handle, so the edit is there
		 ; now do whatever you need to do
		 ExitLoop
	  EndIf
   WEnd
   ControlFocus(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:1]")

   Local $iAnimationWidth = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:1]")
   If @error Then ConsoleWrite("Error reading Animation Width data." & @CRLF)
   Local $iAnimationHeight = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:2]")
   If @error Then ConsoleWrite("Error reading Animation Height data." & @CRLF)
   If $GIFAnimatorDebug = True Then
	  ConsoleWrite("Animation Width: " & $iAnimationWidth & @CRLF)
	  ConsoleWrite("Animation Height: " & $iAnimationHeight & @CRLF)
   EndIf

   ; Change Animation Width / Height
   If $GIFAnimatorDebug = True Then ConsoleWrite("New Width: " & Int($iAnimationWidth / 2) & @CRLF)
   ControlSetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:1]", Int($iAnimationWidth / 2))
   Sleep(2000)
   If $GIFAnimatorDebug = True Then ConsoleWrite("New Height: " & Int($iAnimationHeight / 2) & @CRLF)
   ControlSetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:2]", Int($iAnimationHeight / 2))
   Sleep(2000)
EndFunc

Func GIFAnimatorClose()
   If Not ProcessExists(WinGetProcess(HWnd($GIFAnimatorhWnd))) Then
	  GIFAnimatorQuit()
   EndIf
EndFunc

Func GIFAnimatorQuit()
   Exit
EndFunc

#Region _IsEmpty()
Func _IsEmpty($vVal)
    If IsArray($vVal) And (Eval($vVal) == '') Then
        Return True
    ElseIf ($vVal == '') Or ($vVal == Null) Then
        Return True
    EndIf
    Return False
 EndFunc
#EndRegion _IsEmpty()

Func _GetGIFDuration($sFilePath)
   ; https://autoit.de/thread/46160-animiertes-gif-in-einzelne-frames-splitten/
   _GDIPlus_Startup()
   ; Local $binGif = Binary(FileRead($sFilePath))
   ; Local $hGIFImage = _GDIPlus_BitmapCreateFromMemory($binGif)
   Local $hGIFImage = _GDIPlus_ImageLoadFromFile($sFilePath) ; https://www.autoitscript.com/forum/topic/196504-load-image-handle-of-file/
   Local $iAnimDimCount = _GDIPlus_GIFAnimGetFrameDimensionsCount($hGIFImage)
   Local $tGUID = _GDIPlus_GIFAnimGetFrameDimensionsList($hGIFImage, $iAnimDimCount)
   Local $iAnimFrameCount = _GDIPlus_GIFAnimGetFrameCount($hGIFImage, $tGUID)
   ; _ArrayDisplay(_GDIPlus_GIFAnimGetFrameDelays($hGIFImage, $iAnimFrameCount))
   Local $aDuration =  _GDIPlus_GIFAnimGetFrameDelays($hGIFImage, $iAnimFrameCount)
   _GDIPlus_ImageDispose($hGIFImage)
   _GDIPlus_Shutdown()

   Return $aDuration
EndFunc

Func _FolderExists($sPath)
   If StringInStr(FileGetAttrib($sPath),"D") Then
	  Return True
   Else
	  Return False
   EndIf
EndFunc

Func _WaitChildDialog($hWnd)
	Local $aGIFAnimatorOpenDialogs = _WinChildDialogs($hWnd, 5 * 1000) ; Timeout at 5sec
	If $aGIFAnimatorOpenDialogs[0][0] > 0 Then
		Return $aGIFAnimatorOpenDialogs[1][1] ; Return Child Dialog hWnd
	EndIf

	Return SetError(1, 0, $aGIFAnimatorOpenDialogs[0][0])
EndFunc

Func _Quotes($sString)
   Return Chr(34) & $sString & Chr(34)
EndFunc   ;==>Quotes
