#autoit3wrapper_icon=..\Icons\Gnome\Gnome-Tools-Check-Spelling.ico
#AutoIt3Wrapper_usex64=N
#AutoIt3Wrapper_Res_Fileversion=0.2.5.25
#autoit3wrapper_res_fileversion_autoincrement=Y
#autoit3wrapper_res_description=Auto Rename Utility
#autoit3wrapper_res_legalcopyright=Jon Dunham 2010
#AutoIt3Wrapper_Run_After=copy %out% "\\ftwgen01\this\field services\autoitscripts\autoName.exe"

;~ #RequireAdmin
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <guiedit.au3>
#include <string.au3>

#include "userinfo.au3"
#include "tma2.au3"
#include "_services.au3"

#include <modernMenuRaw.au3>

Opt("TrayAutoPause", 0)
Opt("TrayIconHide", 1)

#Region ### START Koda GUI section ### Form=
$frmMain = GUICreate("autoName", 233, 202, 358, 375);, -1, BitOR($WS_EX_WINDOWEDGE,$WS_EX_COMPOSITED))

$menu = GUICtrlCreateContextMenu()
;~ $menuCheck = GUICtrlCreateMenuItem("Re-check New Name", $menu)
$menuCheck = _GUICtrlCreateODMenuItem("Re-check New Name", $menu, "..\Icons\Gnome\16\stock_spellcheck.ico")
$menuReboot = _GUICtrlCreateODMenuItem("Re-start Computer", $menu, "..\Icons\fugue\Icons\control-power.png")
_GUICtrlCreateODMenuItem("", $menu)
$menuOptions = _GUICtrlCreateODMenuItem("&Options", $menu, "..\Icons\Gnome\16\Preferences-Desktop.ico")

GUICtrlCreateLabel("Computer:", 8, 10, 53, 20)
GUICtrlCreateLabel("New Name:", 8, 34, 60, 20)
GUICtrlCreateLabel("Current User:", 8, 58, 69, 20)
$inName = GUICtrlCreateInput(@ComputerName, 104, 8, 121, 21)
$inNew = GUICtrlCreateInput("", 104, 32, 121, 21)
$inUser = GUICtrlCreateInput("", 104, 56, 121, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_READONLY))
GUICtrlSetCursor(-1, 2)
$btnSet = GUICtrlCreateButton("Rename", 144, 80, 83, 25, BitOR($BS_DEFPUSHBUTTON,$BS_NOTIFY,$BS_FLAT))
GUICtrlSetState(-1, $GUI_DISABLE)
$Edit1 = GUICtrlCreateEdit("", 0, 112, 233, 90, BitOR($ES_READONLY,$WS_VSCROLL))
GUICtrlSetFont(-1, 8, 400, 0, "Verdana")
GUICtrlSetBkColor(-1, 0xA6CAF0)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $pathINI = @AppDataDir & "\TMA2 Tools\autoNameSettings.ini"
if NOT FileExists($pathINI) Then
	DirCreate(@AppDataDir & "\TMA2 Tools")
	IniWriteSection($pathINI, "Options", "InitKey=InitValue")
	ConsoleWriteError("iniwrite error: " & @error & @CRLF)
EndIf

Global $uPass = InputBox("Password Required", "Please enter your network password", "", "*", Default, Default, Default, Default, 60, $frmMain)
if @error Then
	GUIDelete()
	Exit
Else
	$uPass = _StringEncrypt(1, $uPass, @UserName, 5)
EndIf

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $btnSet
			_go()
		Case $inName
			$sName = GUICtrlRead($inName)
			if _checkName($sName) = 1 Then
				GUICtrlSetState($btnSet, $GUI_DISABLE)
				GUICtrlSetData($btnSet, "Working...")
				GUICtrlSetBkColor($inName, 0x7FFF7F)

				if IniRead($pathINI, "Options", "UseRegEx", "") Then
					GUICtrlSetData( _
						$inNew, _
						StringRegExpReplace( _
							$sName, _
							IniRead($pathINI, "Options", "RegExPattern", ""), _
							IniRead($pathINI, "Options", "RegExReplace", "") ))
				Else
					GUICtrlSetData($inNew, StringUpper(StringReplace(GUICtrlRead($inName), "-", "")) & "B")
				EndIf

				_usr()
				_checkNewName()

				GUICtrlSetData($btnSet, "Rename")
			Else
				GUICtrlSetBkColor($inName, 0xFF7F7F)
				GUICtrlSetState($btnSet, $GUI_DISABLE)
			EndIf
		Case $inNew
			_checkNewName()
		Case $menuCheck
			_checkNewName()
		Case $menuOptions
			_guiShowOptions()
		Case $menuReboot
			$sNew = GUICtrlRead($inNew)
			$sOld = GUICtrlRead($inName)
			if NOT Ping($sOld) Then
				_GUICtrlEdit_AppendText($Edit1, "Computer not online." & @CRLF)
				ContinueLoop
			EndIf

			$iTimeout = IniRead($pathINI, "Options", "Timeout", "60")
			$sMessage = IniRead($pathINI, "Options", "Message", "Renaming. Computer will restart automatically.")

			$sMessage = StringReplace($sMessage, "%OldName%", $sOld)
			$sMessage = StringReplace($sMessage, "%NewName%", $sNew)

			$iReturn = ShellExecuteWait("shutdown", "-r -m \\" & $sOld & ' -t ' & $iTimeout & ' -c "' & $sMessage & '" -d U:2:4', "", "open", @SW_HIDE)
			if NOT $iReturn and not @error Then
				_GUICtrlEdit_AppendText($Edit1, "Shutdown executed successfully." & @CRLF)
			ElseIf $iReturn Then
				_GUICtrlEdit_AppendText($Edit1, "Shutdown command failed with exit code " & $iReturn & "." & @CRLF)
			EndIf
	EndSwitch
WEnd

Func _checkNewName()
	GUICtrlSetData($btnSet, "Working...")
	GUICtrlSetState($btnSet, $GUI_DISABLE)
	$sNew = GUICtrlRead($inNew)
	$iCheck = _checkName($sNew)

	if $iCheck = 0 AND NOT @error Then
		_GUICtrlEdit_AppendText($Edit1, "New name " & $sNew & " available." & @CRLF)
		GUICtrlSetBkColor($inNew, 0x7FFF7F)
		GUICtrlSetState($btnSet, $GUI_ENABLE)
	ElseIf $iCheck = 1 Then
		_GUICtrlEdit_AppendText($Edit1, "New name " & $sNew & " in use." & @CRLF)
		GUICtrlSetBkColor($inNew, 0xFF7F7F)
	ElseIf $iCheck = -1 Then
		_GUICtrlEdit_AppendText($Edit1, "New name " & $sNew & " in use (offline)." & @CRLF)
		GUICtrlSetBkColor($inNew, 0xFF7F7F)
	EndIf

	GUICtrlSetData($btnSet, "Rename")
EndFunc

Func _checkName($sName)
	if NOT _computerNameLegal($sName) Then
		Return SetError(1)
	EndIf

	$iPing = Ping($sName)
	if @error = 1 Then
		Return -1
	ElseIf @error >= 2 Then
		; want ret 0 and no error for new name
		Return 0
	EndIf

	Return 1
EndFunc

Func _usr()
	$user = _currentUser(GUICtrlRead($inName))
	$errUser = @error

	if $errUser Then
		_GUICtrlEdit_AppendText($Edit1, "Error " & $errUser & " getting current user." & @CRLF)
		GUICtrlSetData($inUser, "?")
		GUICtrlSetBkColor($inUser, 0xFF3333)
		GUICtrlSetState($btnSet, $GUI_DISABLE)
	Else
		if $user <> "" Then
			$split = StringSplit($user, "\")
			$aAccount = _ADGetUserInfo("SAMAccountName", $split[2])

			$regex = StringRegExp($user, "[a-zA-Z]*\\amh[a-z]{0,3}\d+")
			if $regex Then
				GUICtrlSetData($inUser, $user)
				GUICtrlSetBkColor($inUser, 0xFFFF7F)
			Else
				GUICtrlSetData($inUser, $user)
				GUICtrlSetBkColor($inUser, 0xFF7F7F)
			EndIf

			if IsArray($aAccount) Then
				$sTip = "Full name: " & @TAB & $aAccount[1][1] & @CRLF
				$sTip &= "Department:" & @TAB & $aAccount[1][5] & @CRLF
				$sTip &= "Job title: " & @TAB & $aAccount[1][3] & @CRLF
				if $aAccount[1][7] Then
					$sTip &= "Phone #:   " & @TAB & $aAccount[1][7] & @CRLF
				EndIf
				$sTip &= "Email add: " & @TAB & $aAccount[1][8]

				GUICtrlSetTip($inUser, $sTip, "User Details", 1)
			EndIf
		ElseIf $user = "" Then
			GUICtrlSetData($inUser, "None (Logon screen)")
			GUICtrlSetBkColor($inUser, 0x7FFF7F)
			GUICtrlSetState($btnSet, $GUI_ENABLE)
		EndIf
	EndIf

EndFunc

Func _go()
	if GUICtrlRead($inNew) = "" Then
		GUICtrlSetBkColor($inNew, 0xFF7F7F)
;~ 		_GUICtrlEdit_AppendText($Edit1, "Enter name to contact!" & @CRLF)
		Return
	EndIf

	GUISetCursor(15, 1, $frmMain)
	Local $sOld = GUICtrlRead($inName)
	Local $sNew = GUICtrlRead($inNew)
	Local $iTimeout = IniRead($pathINI, "Options", "Timeout", "60")
	Local $sMessage = IniRead($pathINI, "Options", "Message", "Renaming")

	$sMessage = StringReplace($sMessage, "%OldName%", $sOld)
	$sMessage = StringReplace($sMessage, "%NewName%", $sNew)

	_renameComp($sNew, $sOld, @UserName, _StringEncrypt(0, $uPass, @UserName, 5))
	if @error Then
		$err = @error
		_GUICtrlEdit_AppendText($Edit1, "Error " & $err & " renaming computer." & @CRLF)
		Switch $err
			Case 2224
				_GUICtrlEdit_AppendText($Edit1, "Computer name already exists." & @CRLF)
			Case 1326
				_GUICtrlEdit_AppendText($Edit1, "Unknown user name or bad password." & @CRLF)
				$uPass = InputBox("Password Required", "Please re-enter your network password", "", "*", Default, Default, Default, Default, 60, $frmMain)
			Case 2221
				_GUICtrlEdit_AppendText($Edit1, "The user name could not be found." & @CRLF)
				_GUICtrlEdit_AppendText($Edit1, "The computer was probably renamed and not restarted." & @CRLF)
		EndSwitch
	Else
		_GUICtrlEdit_AppendText($Edit1, "Renamed successfully!" & @CRLF)
		if MsgBox(4132, "Reboot Required", "Reboot computer now?", 60, $frmMain) = 6 Then
			$iReturn = ShellExecuteWait("shutdown", "-r -m \\" & $sOld & ' -t ' & $iTimeout & ' -c "' & $sMessage & '" -d U:2:4', "", "open", @SW_HIDE)
			if NOT $iReturn and not @error Then
				_GUICtrlEdit_AppendText($Edit1, "Shutdown executed successfully." & @CRLF)
			ElseIf $iReturn Then
				_GUICtrlEdit_AppendText($Edit1, "Shutdown command failed with exit code " & $iReturn & @CRLF)
			EndIf
		EndIf
	EndIf

	GUISetCursor()
EndFunc

Func _guiShowOptions()
	$aMainPos = WinGetPos($frmMain)

	#Region ### START Koda GUI section ### Form=C:\Users\TMA2\Documents\Scripts\Forms\frmOptions.kxf
	$frmOptions = GUICreate("Options", 297, 305, $aMainPos[0]-50, $aMainPos[1]+25, BitOR($WS_POPUP,$WS_CAPTION,$WS_SYSMENU,$WS_OVERLAPPED), -1, $frmMain)
	GUICtrlCreateGroup("Restart", 8, 8, 280, 120)
	$inTime = GUICtrlCreateInput("30", 128, 24, 40, 21, BitOR($ES_RIGHT,$ES_AUTOHSCROLL,$ES_NUMBER))
	GUICtrlSetLimit(-1, 2)
	$UpDown1 = GUICtrlCreateUpdown($inTime)
	GUICtrlSetLimit(-1, 99, 0)
	GUICtrlCreateLabel("Time before restart:", 16, 24, 95, 17)
	$inMessage = GUICtrlCreateEdit("", 16, 56, 264, 64, $ES_MULTILINE)
	GUICtrlSetLimit(-1, 127)
	GUICtrlSetData(-1, "Renaming computer from %OldName% to %NewName%. Please save and close any open documents.")
	GUICtrlSetTip(-1, "Old/current name: %OldName%" & @CRLF & "New name: %NewName%", "Variables", 1)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$btnOptOK = GUICtrlCreateButton("&OK", 136, 272, 72, 24, $BS_DEFPUSHBUTTON)
	$btnOptCancel = GUICtrlCreateButton("&Cancel", 216, 272, 72, 24, 0)
	GUICtrlCreateGroup("Renaming", 8, 136, 280, 129)
	$chkRegEx = GUICtrlCreateCheckbox("Auto-fill 'New Name' field w/ regex:", 16, 152, 185, 17)
	$inRegExPattern = GUICtrlCreateInput("(\w+)(?:-*)(\w+)", 64, 176, 177, 21)
	$inRegExReplace = GUICtrlCreateInput("$1WS$2", 64, 200, 177, 21)
	GUICtrlCreateLabel("Pattern:", 16, 176, 41, 17)
	GUICtrlCreateLabel("Replace:", 16, 200, 47, 17)
	$icnPattern = GUICtrlCreateIcon("", 0, 248, 176, 16, 16, BitOR($SS_NOTIFY,$WS_GROUP))
	$icnReplace = GUICtrlCreateIcon("", 0, 248, 200, 16, 16, BitOR($SS_NOTIFY,$WS_GROUP))
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	GUISetState(@SW_DISABLE, $frmMain)
	GUISetState(@SW_SHOW, $frmOptions)

	if IniRead($pathINI, "Options", "UseRegEx", 1) Then
		GUICtrlSetState($chkRegEx, $GUI_CHECKED)
	Else
		GUICtrlSetState($chkRegEx, $GUI_UNCHECKED)
	EndIf

	GUICtrlSetData($inMessage, IniRead($pathINI, "Options", "Message", "Renaming computer from %OldName% to %NewName%. Please save and close any open documents."))
	GUICtrlSetData($inTime, IniRead($pathINI, "Options", "Timeout", "30"))
	GUICtrlSetData($inRegExPattern, IniRead($pathINI, "Options", "RegExPattern", "(\w+)(?:-*)(\w+)"))
	GUICtrlSetData($inRegExReplace, IniRead($pathINI, "Options", "RegExReplace", "$1WS$2"))

	#EndRegion ### END Koda GUI section ###
	Local $bValid1 = True
	Local $bValid2 = True

	While 1
		$msgOpt = GUIGetMsg()
		$patternTemp = GUICtrlRead($inRegExPattern)
		$replaceTemp = GUICtrlRead($inRegExReplace)
		$chkTemp = GUICtrlRead($chkRegEx)

		Switch $msgOpt
			Case $btnOptOK
				if BitAND($chkTemp, $GUI_CHECKED) Then
					if $bValid1 AND $bValid2 Then
						IniWrite($pathINI, "Options", "RegExPattern", GUICtrlRead($inRegExPattern))
						IniWrite($pathINI, "Options", "RegExReplace", GUICtrlRead($inRegExReplace))
					Else
						MsgBox(4096+48, "Warning", "Invalid regular expression or replacement pattern.", -1, $frmOptions)
						ContinueLoop
					EndIf
				EndIf

				IniWrite($pathINI, "Options", "UseRegEx", BitAND(GUICtrlRead($chkRegEx), $GUI_CHECKED))
				IniWrite($pathINI, "Options", "Message", GUICtrlRead($inMessage))
				IniWrite($pathINI, "Options", "Timeout", GUICtrlRead($inTime))
				ExitLoop
			Case $btnOptCancel
				ExitLoop
			Case $GUI_EVENT_CLOSE
				ExitLoop
		EndSwitch

		Sleep(50)

		if GUICtrlRead($chkRegEx) <> $chkTemp Then
			if BitAND(GUICtrlRead($chkRegEx), $GUI_CHECKED) Then
				GUICtrlSetState($inRegExPattern, $GUI_ENABLE)
				GUICtrlSetState($inRegExReplace, $GUI_ENABLE)
			Else
				GUICtrlSetState($inRegExPattern, $GUI_DISABLE)
				GUICtrlSetState($inRegExReplace, $GUI_DISABLE)
			EndIf
		EndIf

		if GUICtrlRead($inRegExPattern) <> $patternTemp Then
			StringRegExp("null", GUICtrlRead($inRegExPattern))
			if @error = 2 Then
				GUICtrlSetImage($icnPattern, "Icons\Gnome\16\Dialog-Error.ico")
				$bValid1 = False
			Else
				GUICtrlSetImage($icnPattern, "Icons\Gnome\16\emblem-default.ico")
				$bValid1 = True
			EndIf
		EndIf

		if GUICtrlRead($inRegExReplace) <> $replaceTemp Then
			if StringRegExp(GUICtrlRead($inRegExReplace), "\$[0-9]") Then
				GUICtrlSetImage($icnReplace, "Icons\Gnome\16\emblem-default.ico")
				$bValid2 = True
			Else
				GUICtrlSetImage($icnReplace, "Icons\Gnome\16\Dialog-Error.ico")
				$bValid2 = False
			EndIf
		EndIf
	WEnd

	GUISetState(@SW_ENABLE, $frmMain)
	GUIDelete($frmOptions)
	Return
EndFunc