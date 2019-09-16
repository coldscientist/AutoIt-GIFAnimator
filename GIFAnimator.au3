#cs ----------------------------------------------------------------------------

 Title: GIFAnimator
 Version: 1.0.0
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

#include <WinAPISysWin.au3>

#include <Misc.au3>

#include <GuiTab.au3>

; _FolderExists
#include "AutoIt-FSOClass\FSOClass.au3"

; IsHungAppWindow
#include <WinAPISysWin.au3>

Opt("MustDeclareVars", 1)

AutoItSetOption("MouseCoordMode", 2)
AutoItSetOption("TrayIconDebug", 1)
;OnAutoItExitRegister("OnAutoItExit")
HotKeySet("{ESCAPE}", "GIFAnimatorQuit")

Global $GIFAnimatorDebug = True

Global $GIFAnimatorPath = Null

If _Singleton(@ScriptName, 1) = 0 Then
   MsgBox($MB_SYSTEMMODAL, "GIFAnimator", "An occurrence of " & @ScriptName & " is already running. Aborting.")
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
		 MsgBox(BitOr($MB_ICONERROR, $MB_SYSTEMMODAL), "GIFAnimator.au3", "GIFAnimator.exe could not be found at " & Quotes($ProgramFilesDir & "\Microsoft GIF Animator\GIFAnimator.exe") & " or " & $ImageComposerAddIn & ". Aborting.")
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
   Run($GIFAnimatorPath)
   $hWnd = WinWait("Microsoft Gif Animator")
   $GIFAnimatorhWnd = String($hWnd)
   If $GIFAnimatorDebug = True Then ConsoleWrite("GIFAnimator.exe hWnd: " & $GIFAnimatorhWnd & @CRLF)
   Global $hToolbars = ControlGetHandle(HWnd($GIFAnimatorhWnd), "", "[CLASS:ToolbarWindow32; INSTANCE:1]")

   ; https://www.autoitscript.com/forum/topic/136041-solved-_filelisttoarray-need-an-explanation/
   For $i = 1 to UBound($aGIF) -1
	  ConsoleWrite( "(" & ($i) & " of " & $aGIF[0] & ") " & $aGIF[$i] & @crlf)

	  ; https://www.autoitscript.com/forum/topic/120516-test-for-window-responsiveness/
	  While 1
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

	  Local $iDuration = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:3]") ; 1 = 10 msec
	  If @error Then
		 ConsoleWrite("Error reading Duration data.")
		 Exit
	  Else
		 ConsoleWrite("Duration: " & $iDuration & @CRLF)
	  EndIf

	  ; https://www.autoitscript.com/autoit3/docs/intro/lang_operators.htm
	  If $iDuration <= 5 Then
		 ContinueLoop ; Do Nothing
	  EndIf

	  ; Select All (CTRL + L)
	  If Not WinActive(HWnd($GIFAnimatorhWnd)) Then WinActivate(HWnd($GIFAnimatorhWnd))
	  _GUICtrlToolbar_ClickIndex($hToolbars, 11)

	  ; Change Duration
	  If $GIFAnimatorDebug = True Then ConsoleWrite("New Duration: " & Int($iDuration / 2) & @CRLF)
	  ControlSetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Edit; INSTANCE:3]", Int($iDuration / 2))

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
		 MsgBox(0, "GIFAnimator", "Could not create backup file in " & Quotes($sBackupFile) & ". Aborting.")
		 Exit
	  Else
		 FileDelete($aGIF[$i])
		 If FileExists($aGIF[$i]) Then
			MsgBox(0, "GIFAnimator", "Could not delete " & Quotes($aGIF[$i]) & " file. Aborting.")
			Exit
		 EndIf
	  EndIf

	  ; Save As (Ctrl + A)
	  If Not WinActive(HWnd($GIFAnimatorhWnd)) Then WinActivate(HWnd($GIFAnimatorhWnd))
	  _GUICtrlToolbar_ClickIndex($hToolbars, 4)
	  Local $shWnd = _WaitSelectDialog()
	  WinActivate($shWnd)
	  Send("{ENTER}")

	  ; Writing
	  ; http://www.autoitscript.com/forum/topic/95905-wait-until-button-text-ok/
	  Local $sReady = Null
	  While 1
		 ; https://www.autoitscript.com/forum/topic/4646-text-with-edit-and-static-classes/
		 $sReady = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Static; INSTANCE:25]")
		 If $GIFAnimatorDebug = True Then ConsoleWrite("Ready (Writing): " & $sReady & @CRLF)
		 Sleep(100)
		 If $sReady = "Ready" Then
			ExitLoop
		 EndIf
	  WEnd
   Next
EndFunc    ;==>GIFAnimator

Func GIFAnimatorOpen($sFilePath)
   If Not WinActive(HWnd($GIFAnimatorhWnd)) Then WinActivate(HWnd($GIFAnimatorhWnd))

   ; New (CTRL + N)
   ; https://www.autoitscript.com/forum/topic/155068-toolsbar-understanding-how-to-access-these-controls/
   _GUICtrlToolbar_ClickIndex($hToolbars, 0)

   ; Open (CTRL + O)
   _GUICtrlToolbar_ClickIndex($hToolbars, 1)

   Local $ohWnd = _WaitSelectDialog()

   ; https://www.autoitscript.com/forum/topic/166908-cant-controll-class32770-windows-7/
   If Not WinActive($ohWnd) Then WinActivate($ohWnd)
   WinWaitActive($ohWnd)
   ControlSetText($ohWnd, "", "[CLASS:Edit; INSTANCE:1]", $sFilePath) ; Set the edit control in Open with some text. The handle returned by WinWait is used for the "title" parameter of ControlSetText.
   ControlSend($ohWnd,"", 1,"{ENTER}")

   ; Reading
   ; http://www.autoitscript.com/forum/topic/95905-wait-until-button-text-ok/
   Local $sReady = Null
   While 1
	  ; https://www.autoitscript.com/forum/topic/4646-text-with-edit-and-static-classes/
	  $sReady = ControlGetText(HWnd($GIFAnimatorhWnd), "", "[CLASS:Static; INSTANCE:8]")
	  If $GIFAnimatorDebug = True Then ConsoleWrite("Ready (Open): " & $sReady & @CRLF)
	  Sleep(100)
	  If $sReady = "Ready" Then
		 ExitLoop
	  EndIf
   WEnd

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

Func _WaitSelectDialog()
   While 1
	  ; Retrieve a list of window handles.
	  Local $aWinList = WinList()

	  ; Run through GUIs to find those associated with the process
	  For $i = 1 To $aWinList[0][0]
		 ; Loop through the array displaying only visable windows with a title.
		 If $aWinList[$i][0] <> "" And BitAND(WinGetState($aWinList[$i][1]), 2) Then
			; If $GIFAnimatorDebug = True Then ConsoleWrite("Title: " & $aWinList[$i][0] & @CRLF & "Handle: " & $aWinList[$i][1] & "PID: " & WinGetProcess($aWinList[$i][1]) & " Parent ID: " & _WinAPI_GetParent ( $aWinList[$i][1] ) & @CRLF)
			; If $GIFAnimatorDebug = True Then ConsoleWrite("If " & _WinAPI_GetParent ( $aWinList[$i][1] ) & " = " & $GIFAnimatorhWnd & " Then" & @CRLF)
			If _WinAPI_GetParent ( $aWinList[$i][1] ) = $GIFAnimatorhWnd Then
			   Return $aWinList[$i][1]
			EndIf
		 EndIf
	  Next
   WEnd
EndFunc   ;==>_WaitSelectDialog

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
			ElseIf _FileExists($CmdLine[$x]) Then
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

		 $GIFOriginalPath = $sDirPath & "\GIF (Original)"
		 If Not _FolderExists($GIFOriginalPath) Then
			If Not DirCreate($GIFOriginalPath) Then
			   MsgBox(0, "GIFAnimator", "Could not create " & Quotes($GIFOriginalPath) & " directory. Aborting.")
			   Exit
			EndIf
		 EndIf
	  EndIf
   EndIf

   If Not IsArray($aGIF) Then
	  MsgBox(0, "GIFAnimator", "No GIF files found at " & Quotes($sDirPath) & ". Aborting.")
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

Func GIFAnimatorQuit()
   Exit
EndFunc

Func Quotes($sString)
   Return Chr(34) & $sString & Chr(34)
EndFunc   ;==>Quotes
