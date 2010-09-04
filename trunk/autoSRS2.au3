#AutoIt3Wrapper_Icon=Icons\BeOS\1265138248_BeOS_MIDI_Video_doc.ico
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Description=Frontend for MKV/AVI ReSample
#AutoIt3Wrapper_Res_Comment=http://rescene.info/resample/
#AutoIt3Wrapper_Res_Field=Product name|autoSRS
#AutoIt3Wrapper_Res_Field=Product version|%AutoItVer%
#AutoIt3Wrapper_Res_LegalCopyright=Jon Dunham 2010
#AutoIt3Wrapper_Res_Fileversion=0.1.1.10
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=y
#AutoIt3Wrapper_Outfile=f:\Software\ReScene\autoSRS.exe

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <guiStatusBar.au3>

#include <guiEdit.au3>
#include <file.au3>

Global $PATH_SRS = @WorkingDir & "\srs.exe"

if NOT FileExists(@WorkingDir & "\srs.exe") Then
	$PATH_SRS = FileOpenDialog("Please select srs executable...", @WorkingDir, "ReSample (srs.exe)")
	if @error Then Exit
EndIf

Opt("GUIOnEventMode", 1)
Opt("TrayAutoPause", 0)
Opt("TrayIconHide", 1)
Opt("GUIResizeMode", $GUI_DOCKALL)
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("autoSRS", 547, 250, 5, 5, $WS_SIZEBOX+$WS_MINIMIZEBOX, $WS_EX_ACCEPTFILES)

$rCreate = GUICtrlCreateRadio("Create ReSample", 8, 8, 105, 17)
GUICtrlSetOnEvent(-1, "rCreate_Click")
$rBackup = GUICtrlCreateRadio("ReCreate Sample", 120, 8, 113, 17)
GUICtrlSetOnEvent(-1, "rBackup_Click")
GUICtrlSetState($rBackup, $GUI_CHECKED)
GUICtrlCreateGroup("File Locations", 8, 8+24, 529, 89)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKRIGHT+$GUI_DOCKHEIGHT)
$in1 = GUICtrlCreateInput("Enter (or drag & drop) .srs", 16, 32+24, 453, 21)
GUICtrlSetColor(-1, 0x666666)
GUICtrlSetResizing(-1, $GUI_DOCKHEIGHT+$GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKRIGHT)
$btn1 = GUICtrlCreateButton("Browse", 476, 32+24, 50, 21, $BS_FLAT+$BS_ICON)
GUICtrlSetImage(-1, "Icons\Gnome-Folder-Open.ico", 0, 0)
GUICtrlSetResizing(-1, $GUI_DOCKSIZE+$GUI_DOCKRIGHT+$GUI_DOCKTOP)
$in2 = GUICtrlCreateInput("Enter (or drag & drop) full .mkv or .avi", 16, 64+24, 453, 21)
GUICtrlSetColor(-1, 0x666666)
GUICtrlSetResizing(-1, $GUI_DOCKHEIGHT+$GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKRIGHT)
$btn2 = GUICtrlCreateButton("Browse", 476, 64+24, 50, 21, $BS_FLAT+$BS_ICON)
GUICtrlSetImage(-1, "Icons\Gnome-Folder-Open.ico", 0, 0)
GUICtrlSetResizing(-1, $GUI_DOCKSIZE+$GUI_DOCKRIGHT+$GUI_DOCKTOP)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$chkVerify = GUICtrlCreateCheckbox("Verify", 8, 104+24, 75, 25)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetTip(-1, "Verify .srs after creation")
$btnOK = GUICtrlCreateButton("OK", 456, 104+24, 75, 25, $BS_DEFPUSHBUTTON)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKSIZE)
$inLog = GUICtrlCreateEdit("", 0, 160, 546, 42, BitOR($ES_AUTOVSCROLL,$ES_READONLY,$ES_WANTRETURN,$WS_VSCROLL,$WS_HSCROLL))
GUICtrlSetFont(-1, 8, 400, 0, "Lucida Console");, 4)
GUICtrlSetBkColor(-1, 0x000000)
GUICtrlSetColor(-1, 0xDDFFDD)
GUICtrlSetCursor(-1, 2)
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)

$ctxFile = GUICtrlCreateMenu("File")
$ctxEdit = GUICtrlCreateMenu("Edit")
$ctxAbout = GUICtrlCreateMenu("About")

$ctxFileSave = GUICtrlCreateMenuItem("Save log", $ctxFile)
GUICtrlSetOnEvent(-1, "ctxLogSave_Click")
$ctxFileInfo = GUICtrlCreateMenuItem("Display SRS info", $ctxFile)
GUICtrlSetOnEvent(-1, "ctxFileInfo_Click")
$ctxFileInfo = GUICtrlCreateMenuItem("Display sample info", $ctxFile)
GUICtrlSetOnEvent(-1, "ctxFileSInfo_Click")
GUICtrlCreateMenuItem("", $ctxFile)
$ctxFileExit = GUICtrlCreateMenuItem("Exit", $ctxFile)
GUICtrlSetOnEvent(-1, "GUI_EVENT_CLOSE")

$ctxEditPref = GUICtrlCreateMenuItem("Preferences", $ctxEdit)
GUICtrlSetOnEvent(-1, "ctxEditPref_Click")

$ctxAboutAbout = GUICtrlCreateMenuItem("About", $ctxAbout)
GUICtrlSetOnEvent(-1, "ctxAboutAbout_Click")

GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_EVENT_CLOSE")
GUISetOnEvent($GUI_EVENT_DROPPED, "GUI_EVENT_DROPPED")
GUICtrlSetOnEvent($btn1, "btn1_Click")
GUICtrlSetOnEvent($btn2, "btn2_Click")
GUICtrlSetOnEvent($btnOK, "btnOK_Click")

GUICtrlSetState($in1, $GUI_DROPACCEPTED)
GUICtrlSetState($in2, $GUI_DROPACCEPTED)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

#Region ### START Koda GUI section ### Form=C:\Documents and Settings\Jon\My Documents\Work\AI3 scripts\autoSRSPrefs.kxf
$hPrefs = GUICreate("Preferences", 300, 330, @DesktopWidth/2, @DesktopHeight/2)
GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_EVENT_CLOSE_PREFS", $hPrefs)
$btnPrefCancel = GUICtrlCreateButton("&Cancel", 216, 296, 75, 25, $WS_GROUP)
GUICtrlSetOnEvent(-1, "btnPrefCancel_Click")
$btnPrefOK = GUICtrlCreateButton("&OK", 136, 296, 75, 25, $BS_DEFPUSHBUTTON+$WS_GROUP)
GUICtrlSetOnEvent(-1, "btnPrefOK_Click")
GUICtrlCreateGroup("Create", 8, 8, 284, 161)
$chkAllSRS = GUICtrlCreateCheckbox("Use one directory for created .srs files", 16, 112, 201, 17)
$inAllSRS = GUICtrlCreateInput("", 16, 136, 185, 21)
$btnAllSRS = GUICtrlCreateButton("Browse", 208, 136, 75, 21, $WS_GROUP)
GUICtrlSetOnEvent(-1, "btnAllSRS_Click")
$chkB = GUICtrlCreateCheckbox("Support for big files (samples > 2GB)", 16, 28, 201, 17)
$RadioD = GUICtrlCreateRadio("Sample directory name > .srs file name", 16, 56, 209, 17)
$RadioDD = GUICtrlCreateRadio("Parent directory name > .srs file name", 16, 72, 201, 17)
$RadioDDD = GUICtrlCreateRadio("Parent directory name / place in parent directory", 16, 88, 265, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlCreateGroup("Backup", 8, 176, 284, 49)
$chkSampleDir = GUICtrlCreateCheckbox("Put created sample in \Sample directory", 16, 196, 217, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)
GUICtrlCreateGroup("srs.exe Path", 8, 232, 284, 49)
$in1exe = GUICtrlCreateInput("", 16, 252, 185, 21)
$btn1exe = GUICtrlCreateButton("Browse", 208, 252, 75, 21, $WS_GROUP)
GUICtrlSetOnEvent(-1, "btn1exe_Click")
GUICtrlCreateGroup("", -99, -99, 1, 1)
#EndRegion ### END Koda GUI section ###

#Region ### START Koda GUI section ### Form=C:\Documents and Settings\Jon\My Documents\Work\AI3 scripts\About.kxf
$hAbout = GUICreate("About", 321, 178, 302, 218)
GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_EVENT_CLOSE_ABOUT", $hAbout)
GUICtrlCreateGroup("", 8, 8, 305, 129)
GUICtrlCreateLabel("autoSRS", 16, 24, 65, 20, $WS_GROUP)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$lblVersion = GUICtrlCreateLabel("Version", 16, 48, 39, 17, $WS_GROUP)
$lblCopy2 = GUICtrlCreateLabel("ReSample © ReScene.com 2008-2009", 16, 104, 190, 17, $WS_GROUP)
GUICtrlSetOnEvent(-1, "lblCopy2_Click")
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlSetCursor (-1, 0)
$lblCopy1 = GUICtrlCreateLabel("Frontend © Jon Dunham 2010", 16, 80, 148, 17, $WS_GROUP)
GUICtrlSetOnEvent(-1, "lblCopy1_Click")
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlSetCursor (-1, 0)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$btnAboutOK = GUICtrlCreateButton("&OK", 123, 144, 75, 25, $BS_DEFPUSHBUTTON)
GUICtrlSetOnEvent(-1, "btnAboutOK_Click")
#EndRegion ### END Koda GUI section ###

#Region Setup
$PATH_INI = @TempDir & "\autoSRS.ini"
if NOT FileExists($PATH_INI) Then
	IniWriteSection($PATH_INI, "Options", "LastPosX=10" & @LF & "LastPosY=10" & @LF & "LastWidth=500" & @LF & "LastHeight=350" & @LF & "LastDirSRS=" & @MyDocumentsDir & @LF & "LastDirVideo" & @MyDocumentsDir & @LF & "SRSexe=" & $PATH_SRS)
	IniWriteSection($PATH_INI, "Preferences", "B=4" & @LF & "DDD=3" & @LF & "chkAllSRSDir=4" & @LF & "AllSRSDir=" & @LF & "PutInSample=1")
Else
	$iLastX = IniRead($PATH_INI, "Options", "LastPosX", 10)
	$iLastY = IniRead($PATH_INI, "Options", "LastPosY", 10)
	$iLastW = IniRead($PATH_INI, "Options", "LastWidth", 500)
	$iLastH = IniRead($PATH_INI, "Options", "LastHeight", 250)
	WinMove($Form1, "", $iLastX, $iLastY, $iLastW, $iLastH)
EndIf

_updatePrefs()

$defMsg1 = "Enter (or drag & drop) .srs"
$defMsg2 = "Enter (or drag & drop) full .mkv or .avi"

#EndRegion Setup

While 1
	#Region Web 2.0 lookin' form colors
	$tempFocus = ControlGetFocus($Form1)
	Sleep(100)
	if $tempFocus = ControlGetFocus($Form1) Then ContinueLoop
	Switch ControlGetFocus($Form1)
		Case "Edit1"
			GUICtrlSetColor($in1, 0x000000)
			GUICtrlSetBkColor($in1, 0x33FF99)
			if StringInStr(GUICtrlRead($in1), $defMsg1) Then GUICtrlSetData($in1, "")

			if GUICtrlRead($in2) = "" Then GUICtrlSetData($in2, $defMsg2)
			GUICtrlSetColor($in2, 0x666666)
			GUICtrlSetBkColor($in2, Default)
		Case "Edit2"
			GUICtrlSetColor($in2, 0x000000)
			GUICtrlSetBkColor($in2, 0x33FF99)
			if StringInStr(GUICtrlRead($in2), $defMsg2) Then GUICtrlSetData($in2, "")

			if GUICtrlRead($in1) = "" Then GUICtrlSetData($in1, $defMsg1)
			GUICtrlSetColor($in1, 0x666666)
			GUICtrlSetBkColor($in1, Default)
		Case Else
			GUICtrlSetBkColor($in1, Default)
			GUICtrlSetBkColor($in2, Default)
			if GUICtrlRead($in2) = "" Then GUICtrlSetData($in2, $defMsg2)
			if GUICtrlRead($in1) = "" Then GUICtrlSetData($in1, $defMsg1)
			if StringInStr(GUICtrlRead($in1), "Enter path to") Then GUICtrlSetColor($in1, 0x666666)
			if StringInStr(GUICtrlRead($in2), "Enter path to") Then GUICtrlSetColor($in2, 0x666666)
		EndSwitch
	#EndRegion
WEnd

Func btn1_Click()
	$lastDir = IniRead($PATH_INI, "Options", "LastDirSRS", @MyDocumentsDir)
	if $lastDir Then
		$sTemp = FileOpenDialog("Open", $lastDir, "ReSample (*.srs)", Default, "", $Form1)
	Else
		$sTemp = FileOpenDialog("Open", @MyDocumentsDir, "ReSample (*.srs)", Default, "", $Form1)
	EndIf
	if @error Then Return

	IniWrite($PATH_INI, "Options", "LastDirSRS", $sTemp)
	GUICtrlSetData($in1, $sTemp)
EndFunc

Func btn2_Click()
	$lastDir = IniRead($PATH_INI, "Options", "LastDirVideo", @MyDocumentsDir)
	if $lastDir Then
		$sTemp = FileOpenDialog("Open", $lastDir, "Matroska (*.mkv)|Audio-Video Interleave (*.avi)", Default, "", $Form1)
	Else
		$sTemp = FileOpenDialog("Open", @MyDocumentsDir, "Matroska (*.mkv)|Audio-Video Interleave (*.avi)", Default, "", $Form1)
	EndIf
	if @error Then Return

	IniWrite($PATH_INI, "Options", "LastDirVideo", $sTemp)
	GUICtrlSetData($in2, $sTemp)
EndFunc

Func btnOK_Click()
	if GUICtrlRead($rBackup) = $GUI_CHECKED Then
		_SRSBackup()
	ElseIf GUICtrlRead($rCreate) = $GUI_CHECKED Then
		_SRSCreate()
	EndIf
EndFunc

Func rCreate_Click()
	$defMsg1 = "Enter (or drag & drop) sample .mkv or .avi"
	$defMsg2 = "Enter (or drag & drop) full .mkv or .avi to check against"
	GUICtrlSetData($in1, $defMsg1)
	GUICtrlSetData($in2, $defMsg2)
	GUICtrlSetState($chkVerify, $GUI_ENABLE)
EndFunc

Func rBackup_Click()
	$defMsg1 = "Enter (or drag & drop) .srs"
	$defMsg2 = "Enter (or drag & drop) full .mkv or .avi"
	GUICtrlSetData($in1, $defMsg1)
	GUICtrlSetData($in2, $defMsg2)
	GUICtrlSetState($chkVerify, $GUI_DISABLE)
EndFunc

Func _SRSBackup()
	Local $sArgs = "", $sOutFile
	If NOT FileExists(GUICtrlRead($in1)) Then
		_GUICtrlEdit_AppendText($inLog, "Invalid .srs path!" & @CRLF)
		Return
	ElseIf NOT FileExists(GUICtrlRead($in2)) Then
		_GUICtrlEdit_AppendText($inLog, "Invalid .mkv path!" & @CRLF)
		Return
	EndIf

	if NOT StringInStr(GUICtrlRead($in2), @WorkingDir) Then
		FileChangeDir(StringLeft(GUICtrlRead($in2), StringInStr(GUICtrlRead($in2), "\", 0, -1)))
	EndIf

	$sArgs = '"' & GUICtrlRead($in1) & '"'
	$sArgs &=  ' "' & GUICtrlRead($in2) & '"'

	$hPID = Run($PATH_SRS & " " & $sArgs, @WorkingDir, @SW_HIDE, 0x8)
	$hTimer = TimerInit()
	_GUICtrlEdit_AppendText($inLog, "---------- MKV ReSample started" & @CRLF)
	_GUICtrlEdit_AppendText($inLog, @TAB & "Working directory: " & @WorkingDir & @CRLF)
	While ProcessExists($hPID)
		$sOut = StdoutRead($hPID)
		if StringInStr($sOut, Chr(8)) Then
			$sOut = StringReplace($sOut, Chr(8), "")
			$sOut = StringReplace($sOut, "\", "")
			$sOut = StringReplace($sOut, "-", "")
			$sOut = StringReplace($sOut, "/", "")
			$sOut = StringReplace($sOut, "|", "")
		EndIf
		if StringInStr($sOut, "Successfully rebuilt") Then _
			$sOutFile = StringReplace($sOut, "Successfully rebuilt sample: ", @WorkingDir & "\")
		_GUICtrlEdit_AppendText($inLog, $sOut)
;~ 		Sleep(10)
	WEnd
	_GUICtrlEdit_AppendText($inLog, StdoutRead($hPID))
	;Successfully rebuilt sample: imbtxvidsnyksample.avi
	if BitAND(GUICtrlGetState($chkSampleDir), $GUI_CHECKED) Then
		_GUICtrlEdit_AppendText($inLog, "Exists: " & $sOutFile & ": " & FileExists($sOutFile) & @CRLF)
	EndIf

	_GUICtrlEdit_AppendText($inLog, "---------- MKV ReSample done after " & Round(TimerDiff($hTimer)/1000, 2) & " seconds." & @CRLF)
EndFunc

Func _SRSCreate()
	if NOT FileExists(GUICtrlRead($in2)) Then
		_GUICtrlEdit_AppendText($inLog, "Invalid .mkv/.avi path!" & @CRLF)
		Return
	EndIf

	$sArgs =  '"' & GUICtrlRead($in2) & '" -ddd'

	$hPID = Run($PATH_SRS & " " & $sArgs, @WorkingDir, @SW_HIDE, 0x8)
	$hTimer = TimerInit()
	_GUICtrlEdit_AppendText($inLog, "---------- MKV ReSample started..." & @CRLF)
	_GUICtrlEdit_AppendText($inLog, @TAB & "Working directory: " & @WorkingDir & @CRLF)
	While ProcessExists($hPID)
		$sOut = StdoutRead($hPID)
		$sOut = StringReplace($sOut, Chr(8), "")
		$sOut = StringReplace($sOut, "\", "")
		$sOut = StringReplace($sOut, "-", "")
		$sOut = StringReplace($sOut, "/", "")
		$sOut = StringReplace($sOut, "|", "")
		_GUICtrlEdit_AppendText($inLog, $sOut)
;~ 		Sleep(10)
	WEnd
	_GUICtrlEdit_AppendText($inLog, StdoutRead($hPID))
	_GUICtrlEdit_AppendText($inLog, "---------- MKV ReSample done after " & Round(TimerDiff($hTimer)/1000, 2) & " seconds." & @CRLF)
EndFunc

Func ctxFileInfo_Click()
	Local $sArgs = ""
	If NOT FileExists(GUICtrlRead($in1)) Then
		_GUICtrlEdit_AppendText($inLog, "Invalid .srs path!" & @CRLF)
		Return
	EndIf

	$sArgs = '"' & GUICtrlRead($in1) & '" -l'

	$hPID = Run($PATH_SRS & " " & $sArgs, @WorkingDir, @SW_HIDE, 0x8)
	While ProcessExists($hPID)
		$sOut = StdoutRead($hPID)
		$sOut = StringReplace($sOut, Chr(8), "")
		$sOut = StringReplace($sOut, "\", "")
		$sOut = StringReplace($sOut, "-", "")
		$sOut = StringReplace($sOut, "/", "")
		$sOut = StringReplace($sOut, "|", "")
		_GUICtrlEdit_AppendText($inLog, $sOut)
	WEnd
	_GUICtrlEdit_AppendText($inLog, StdoutRead($hPID))
	_GUICtrlEdit_AppendText($inLog, " - Done" & @CRLF)
EndFunc

Func ctxFileSInfo_Click()
	Local $sArgs = ""
	If NOT FileExists(GUICtrlRead($in2)) Then
		_GUICtrlEdit_AppendText($inLog, "Invalid .mkv path!" & @CRLF)
		Return
	EndIf

	$sArgs = '"' & GUICtrlRead($in2) & '" -i'

	$aMKVSplit = StringSplit(GUICtrlRead($in2), "\")

	_GUICtrlEdit_AppendText($inLog, "Info for sample: " & @CRLF)
	_GUICtrlEdit_AppendText($inLog, @TAB & $aMKVSplit[$aMKVSplit[0]] & @CRLF)
	$hPID = Run($PATH_SRS & " " & $sArgs, @WorkingDir, @SW_HIDE, 0x8)
	While ProcessExists($hPID)
		$sOut = StdoutRead($hPID)
		$sOut = StringReplace($sOut, Chr(8), "")
		$sOut = StringReplace($sOut, "\", "")
		$sOut = StringReplace($sOut, "/", "")
		$sOut = StringReplace($sOut, "|", "")
		_GUICtrlEdit_AppendText($inLog, $sOut)
		ConsoleWrite($sOut & " | ")
	WEnd
	_GUICtrlEdit_AppendText($inLog, StdoutRead($hPID))
	_GUICtrlEdit_AppendText($inLog, " - Done" & @CRLF)
EndFunc

Func ctxLogSave_Click()
	$pathSave = FileSaveDialog("Select filename to save to", @MyDocumentsDir, "All (*.*)|Text (*.txt)|Log (*.log)")
	if @error Then Return
	FileWrite($pathSave, GUICtrlRead($inLog))
EndFunc

Func ctxAboutAbout_Click()
	GUISetState(@SW_DISABLE, $Form1)
	GUISetState(@SW_SHOW, $hAbout)
EndFunc

Func ctxEditPref_Click()
	GUISetState(@SW_DISABLE, $Form1)
	GUISetState(@SW_SHOW, $hPrefs)
EndFunc

Func btnPrefOK_Click()
	_savePrefs()
	GUISetState(@SW_ENABLE, $Form1)
	GUISetState(@SW_HIDE, $hPrefs)
EndFunc

Func btnPrefCancel_Click()
	_updatePrefs()
	GUISetState(@SW_ENABLE, $Form1)
	GUISetState(@SW_HIDE, $hPrefs)
EndFunc

Func btnAllSRS_Click()
	$sDir = FileSelectFolder("Select folder for created .srs files", "", 7, "", $hPrefs)
	if @error Then Return 0
	GUICtrlSetData($inAllSRS, $sDir)
EndFunc

Func btn1exe_Click()
	$PATH_SRS = FileOpenDialog("Please select srs executable...", @WorkingDir, "ReSample (srs.exe)")
	if @error Then Return 0
	GUICtrlSetData($in1exe, $PATH_SRS)
EndFunc

Func lblCopy1_Click()
	ShellExecute("mailto:jbomb22@gmail.com")
EndFunc

Func lblCopy2_Click()
	ShellExecute("http://www.rescene.com")
EndFunc

Func _savePrefs()
	Local $sINIDataPrefs, $sINIDataOpt

	$sINIDataPrefs &= "B=" & GUICtrlGetState($chkB) & @LF
	if GUICtrlGetState($RadioD) = $GUI_CHECKED Then
		$sINIDataPrefs &= "DDD=" & 1 & @LF
	ElseIf GUICtrlGetState($RadioDD) = $GUI_CHECKED Then
		$sINIDataPrefs &= "DDD=" & 2 & @LF
	ElseIf GUICtrlGetState($RadioDDD) = $GUI_CHECKED Then
		$sINIDataPrefs &= "DDD=" & 3 & @LF
	EndIf
	$sINIDataPrefs &= "chkAllSRSDir=" & GUICtrlGetState($chkAllSRS) & @LF
	$sINIDataPrefs &= "AllSRSDir=" & GUICtrlRead($inAllSRS) & @LF
	$sINIDataPrefs &= "PutInSample=" & GUICtrlGetState($chkSampleDir) & @LF

	IniWriteSection($PATH_INI, "Preferences", $sINIDataPrefs)

	IniWrite($PATH_INI, "Options", "SRSexe", GUICtrlRead($in1exe))
EndFunc

Func _updatePrefs()
	GUICtrlSetState($chkB, IniRead($PATH_INI, "Preferences", "B", ""))
	Switch IniRead($PATH_INI, "Preferences", "DDD", 3)
		Case 1
			GUICtrlSetState($RadioD, $GUI_CHECKED)
		Case 2
			GUICtrlSetState($RadioDD, $GUI_CHECKED)
		Case 3
			GUICtrlSetState($RadioDDD, $GUI_CHECKED)
	EndSwitch
	GUICtrlSetState($chkAllSRS, IniRead($PATH_INI, "Preferences", "chkAllSRSDir", ""))
	GUICtrlSetData($inAllSRS, IniRead($PATH_INI, "Preferences", "AllSRSDir", ""))
	GUICtrlSetState($chkSampleDir, IniRead($PATH_INI, "Preferences", "PutInSample", ""))
	GUICtrlSetData($in1exe, IniRead($PATH_INI, "Options", "SRSexe", $PATH_SRS))
EndFunc

Func GUI_EVENT_CLOSE_PREFS()
	$temp = MsgBox(35, "Warning", "Save changes?")

	If $temp = 2 OR $temp = -1 Then
		Return
	ElseIf $temp = 6 Then
		_savePrefs()
	ElseIf $temp = 7 Then
		_updatePrefs()
	EndIf

	GUISetState(@SW_ENABLE, $Form1)
	GUISetState(@SW_HIDE, $hPrefs)
EndFunc

Func btnAboutOK_Click()
	GUI_EVENT_CLOSE_ABOUT()
EndFunc

Func GUI_EVENT_CLOSE_ABOUT()
	GUISetState(@SW_HIDE, $hAbout)
	GUISetState(@SW_ENABLE, $Form1)
EndFunc

Func GUI_EVENT_CLOSE()
	$aPos = WinGetPos($Form1)
	IniWrite($PATH_INI, "Options", "LastPosX", $aPos[0])
	IniWrite($PATH_INI, "Options", "LastPosY", $aPos[1])
	IniWrite($PATH_INI, "Options", "LastWidth", $aPos[2])
	IniWrite($PATH_INI, "Options", "LastHeight", $aPos[3])
	GUIDelete()
	Exit
EndFunc

Func GUI_EVENT_DROPPED()
	Switch @GUI_DropId
		Case $in1
			GUICtrlSetData($in1, @GUI_DragFile)
		Case $in2
			GUICtrlSetData($in2, @GUI_DragFile)
			$sWD = StringLeft(@GUI_DragFile, StringInStr(@GUI_DragFile, "\", 0, -1))
			FileChangeDir($sWD)
	EndSwitch
EndFunc
