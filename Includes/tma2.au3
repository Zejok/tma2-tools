#include-once
#include <string.au3>
#include <ie.au3>
#include <winapi.au3>

if NOT IsDeclared("debug") Then Global $debug = 0

Func _stringHasUpper($sString)
	$aSplit = StringSplit($sString, "")

	For $i=1 To $aSplit[0]
		if StringIsAlpha($aSplit[$i]) Then
			if StringIsUpper($aSplit[$i]) Then Return 1
		EndIf
	Next
	Return 0
EndFunc

Func _stringHasLower($sString)
	$aSplit = StringSplit($sString, "")

	For $i=1 To $aSplit[0]
		if StringIsAlpha($aSplit[$i]) Then
			if StringIsLower($aSplit[$i]) Then Return 1
		EndIf
	Next
	Return 0
EndFunc

Func _stringConvertUUID($UUID)
	Return StringRegExpReplace($UUID, "([A-F0-9]{2})([A-F0-9]{2})([A-F0-9]{2})([A-F0-9]{2})-([A-F0-9]{2})([A-F0-9]{2})-([A-F0-9]{2})([A-F0-9]{2})", "$4$3$2$1-$6$5-$8$7")
EndFunc

Func _progressAscii($iPercent, $bText = False, $iLength = 20, $cOn = "|", $cOff = "-")
	; [||||||||||----------]

	$sReturn = "["
	$iPercent = $iPercent/100
	$iPercent2 = 1-$iPercent

	$sReturn &= _StringRepeat($cOn, Round($iLength*$iPercent, 0))
	$sReturn &= _StringRepeat($cOff, Round($iLength*$iPercent2, 0))
	$sReturn &= "]"

	if $bText Then $sReturn = _StringInsert($sReturn, "%" & Round($iPercent*100, 0), Round($iLength/2, 0))

	Return $sReturn
EndFunc

#cs
Func _isIP($IP) ; working: maybe use StringRegExp
	$arrIP = StringSplit($IP, ".")

	if $arrIP[0] = 4 Then
		for $i=1 To $arrIP[0]
			if StringIsDigit($arrIP[$i]) Then
				if Number($arrIP[$i]) > 255 OR $arrIP[$i] < 0 Then
					Return SetError(2, $i)
				EndIf
			Else
				Return SetError(3, $i)
			EndIf
		Next
	Else
		Return SetError(1)
	EndIf

	; passed all tests
	Return 1
EndFunc
#ce

Func _isIP($strIP)
	Return StringRegExp($strIP, "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b")
EndFunc

Func _ipGetLocation($sIP)
	_isIP($sIP)
	if @error Then
		Return SetError(@error, @extended)
	EndIf

	$aIPSplit = StringSplit($sIP, ".")

	Switch Number($aIPSplit[1])
	Case 167
		$sReturn = "THFW (Fort Worth)"
	Case 172
		$sReturn = "THAM (Arlington Memorial)"
	Case 10
		Switch Number($aIPSplit[2])
		Case 200
			$sReturn = "Area 0"
		Case 201
			$sReturn = "THFW (Fort Worth)"
		Case 202
			$sReturn = "PHD (Dallas)"
		Case 208
			$sReturn = "THHEB (Hurst/Euless/Bedford)"
		Case 209
			$sReturn = "HSW (Southwest Fort Worth)"
		Case 210
			$sReturn = "THNW (Northwest/Azle)"
		Case 211
			$sReturn = "THS (Stephenville)"
		Case 212
			$sReturn = "THC (Cleburne)"
		Case 213
			$sReturn = "Western"
		Case 214
			$sReturn = "Fort Worth Remote Site"
		Case 215
			$sReturn = "THAM (Arlington Memorial)"
		Case 224
			$sReturn = "THP (Plano)"
		Case 225
			$sReturn = "THA (Allen)"
		Case 226
			$sReturn = "THK (Kaufman)"
		Case 227
			$sReturn = "THW (Winnsboro)"
		Case 228
			$sReturn = "PVN"
		Case 229
			$sReturn = "Dallas Remote Site"
		Case 240
			$sReturn = "THR HQ (Arlington)"
		Case Else
			$sReturn = "Unknown"
		EndSwitch
	Case Else
		$sReturn = "Unknown"
	EndSwitch

	Return $sReturn
EndFunc

#cs old SID-based HKU find
Func _defPrinterInfo($SID, $compName = ".")
	$defPrintString = RegRead("\\" & $compName & "\HKU\" & $SID & "\Software\Microsoft\Windows NT\CurrentVersion\Windows", "Device")

	$info = StringSplit($defPrintString, ",", 2)
	$defPrinter = $info[0]

	Return $defPrinter
EndFunc
#ce

Func _defPrinterInfo($compName = ".")
	Local $info[9], $defaultFound = False

	Local $objWMIService = ObjGet("winmgmts:\\" & $compName & "\root\cimv2")
	if NOT IsObj($objWMIService) Then Return SetError(1, 0, $info)

	Local $colPrt = $objWMIService.ExecQuery("SELECT DeviceID, Status, ExtendedPrinterStatus, " & _
		"PrinterState, Attributes, Network, ServerName, PortName, ShareName FROM Win32_Printer WHERE Default=True" )

	if $colPrt.Count = 0 Then
		Return SetError(2, 0, $info)
	EndIf

	For $objPrt in $colPrt
		$info[0] = $objPrt.DeviceID ; str
		$info[1] = $objPrt.Status ; str
		$info[2] = $objPrt.ExtendedPrinterStatus ; uint
		$info[3] = $objPrt.PrinterState ; uint
		$info[4] = $objPrt.Attributes ; uint
		$info[5] = $objPrt.Network ; bool
		$info[6] = $objPrt.ServerName ; str
		$info[7] = $objPrt.PortName ; str
		$info[8] = $objPrt.ShareName ; str
		$defaultFound = True
	Next

	if $defaultFound Then
		Return $info
	Else
		Return SetError(3, 0, $info)
	EndIf
EndFunc

Func _defPrinterInfo2($compName = ".")
	Dim $i = 1, $sInfo = ""

	While 1
		;S-1-5-21-143154926-1409819978-1803697834-57724
		$sKey = RegEnumKey("\\" & $compName & "\HKU\", $i)
		if @error Then ExitLoop
		$i += 1
		if StringInStr($sKey, "Classes") Then ContinueLoop

		if StringLen($sKey) >= 45 Then
			$sInfo = RegRead("\\" & $compName & "\HKU\" & $sKey & "\Software\Microsoft\Windows NT\CurrentVersion\Windows", "Device")
			ExitLoop
		EndIf

	WEnd

	if $sInfo = "" Then Return SetError(1)

	$sTemp = StringSplit($sInfo, ",", 2)

	Return $sTemp[0]
EndFunc

Func _defPrinterSet($sPrinter)
	Local $sPrinters = ""
	Local $oWSH = ObjCreate("Wscript.Network")
	if @error OR NOT IsObj($oWSH) Then
		MsgBox(4096+48, "Error", "Couldn't create Wscript.Network object. Set default printer manually (" & $sPrinter & ")")
		Return SetError(1)
	EndIf

	$oPrinters = $oWSH.EnumPrinterConnections

	For $i=1 To $oPrinters.Count-1 Step 2
		$sPrinters &= $oPrinters.Item($i) & "|"
	Next

	if NOT StringInStr($sPrinters, $sPrinter) Then
		$oWSH.AddWindowsPrinterConnection($sPrinter)
		if @error Then Return SetError(2)
	EndIf

	$oWSH.SetDefaultPrinter($sPrinter)
	if @error Then Return SetError(3)

	Return 1
EndFunc

func _computerNameLegal($compName)
	dim $i
	Const $illegalChars = StringToASCIIArray( "`~!@#$ ^&*()=+[]{}\|;:',<>/?""" )

	for $i = 0 to UBound($illegalChars)-1
		if StringInStr($compName, Chr($illegalChars[$i])) Then
			; return 0 if computer name contains bad characters
			Return 0
		EndIf

	Next

	; return 1 if computer name is OK, in keeping with 1=success/0=failure autoit function return codes
	Return 1
EndFunc

Func _currentUser($compName)
	dim $objWMIService = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & $compName & "\root\cimv2")

	if NOT IsObj( $objWMIService ) Then
		Return SetError(2)
	EndIf

	$colCS = $objWMIService.InstancesOf('Win32_ComputerSystem')
	for $objCS in $colCS
		Return $objCS.UserName
	Next

EndFunc

Func _listLocalAccounts($compName)
	Dim $oWMIService = ObjGet("winmgmts:\\" & $compName & "\root\cimv2")

	if NOT IsObj( $oWMIService ) Then
		SetError(1)
		Return 0
	EndIf

	$cAccounts = $oWMIService.ExecQuery("SELECT * FROM Win32_UserAccount WHERE LocalAccount=FALSE AND Disabled=FALSE AND Lockout=FALSE")

	For $oAccount in $cAccounts
		ConsoleWrite("-( " & $oAccount.Domain & "\" & $oAccount.Name & " )--------------" & @CRLF)
		ConsoleWrite(@TAB & "Full Name:" & @TAB & $oAccount.FullName & " (" & $oAccount.AccountType & ")" & @CRLF)
		ConsoleWrite(@TAB & "Description:" & @TAB & $oAccount.Description & @CRLF)
	Next
EndFunc

func _getDN($compName, $onlyOU = 1)
	Dim $wshNetwork = ObjCreate("Wscript.Network")
	Dim $objTrans, $objDomain
	Const $ADS_NAME_TYPE_1779 = 1
    Const $ADS_NAME_INITTYPE_GC = 3
    Const $ADS_NAME_TYPE_NT4 = 3

    $objTranslate = ObjCreate("NameTranslate")
	if @error Then
		Return SetError(@Error)
	EndIf
    $objDomain = ObjGet("LDAP://rootDse")
	if @error Then
		Return SetError(@Error)
	EndIf

    $objTranslate.Init($ADS_NAME_INITTYPE_GC, "")

	if $wshNetwork.UserDomain = "" Then
		$objTranslate.Set($ADS_NAME_TYPE_NT4, "TEXAS\" & $compName & "$")
	Else

		$objTranslate.Set($ADS_NAME_TYPE_NT4, $wshNetwork.UserDomain & "\" & $compName & "$")
	EndIf
    $compDN = $objTranslate.Get($ADS_NAME_TYPE_1779)
	if $compDN = "" Then Return SetError(2)
    ;Set DN to upper Case
    $compDN = StringUpper($compDN)

	if $onlyOU Then
		$arrDN = StringSplit($compDN, ",", 2 )
		$compDN = StringTrimLeft($arrDN[1], 3)
	EndIf

	Return $compDN
EndFunc

Func _getHPPowerRecovery($compName = ".")
	$oWMI = ObjGet("winmgmts:\\" & $compName & "\root\hp\instrumentedbios")
	if NOT IsObj($oWMI) Then Return SetError(1)

	$colSetting = $oWMI.execquery('Select Name,Value from HP_BIOSSetting Where Name="After Power Loss"')

	for $objSetting in $colSetting
		Return $objSetting.Value
	Next
EndFunc

Func _setHPPowerRecovery($compName = ".", $sSetting = "On")
	$oWMI = ObjGet("winmgmts:\\" & $compName & "\root\hp\instrumentedbios")
	if NOT IsObj($oWMI) Then Return SetError(1)

	Dim $iReturn
	$colSetting = $oWMI.ExecQuery('Select * from HP_BIOSSettingInterface')

	For $objSetting in $colSetting
		$objSetting.SetBIOSSetting($iReturn, "After Power Loss", $sSetting)
	Next

	if $iReturn = 0 Then
		Return 1
	Else
		Return SetError(2, $iReturn)
	EndIf
EndFunc

Func _appUninstallFind($search, $strComp = ".", $findAll = False, $bRegEx = False)
	; TODO:	implement option to return 2D array of all matches rather than 1
	;		add 64bit support
	Dim $i = 1, $iFound = 0, $strKey, $strName, $appFound

	if $findAll Then
		Dim $arrReturn[1][3]
	Else
		Dim $arrReturn[3]
	EndIf

	While 1
		$strKey = RegEnumKey( "\\" & $strComp & "\HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall", $i )
		if @error Then Return SetError(1)

		$strName = RegRead( "\\" & $strComp & "\HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & $strKey, "DisplayName" )

		if @error Then
			$i += 1
			ContinueLoop
		EndIf

		if $bRegEx Then
			$appFound = StringRegExp($strName, $search)
		Else
			$appFound = StringInStr($strName, $search)
		EndIf

		if $appFound Then
			$arrReturn[0] = $strName
			$arrReturn[1] = RegRead( "\\" & $strComp & "\HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & $strKey, "DisplayVersion" )
			$arrReturn[2] = RegRead( "\\" & $strComp & "\HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & $strKey, "UninstallString" )
			if StringInStr($arrReturn[2], "msiexec") Then
				$arrReturn[2] = StringReplace($arrReturn[2], "/I", "/x") & " /passive"
			EndIf

			if NOT $findAll Then
				ExitLoop
			EndIf
		EndIf
		$i += 1
	WEnd

	Return $arrReturn
EndFunc

Func _delTree($dir, ByRef $files, $progress = 1, $delTop = 0, $topLevel = 1, $totalFiles = 0)
	; Returns # of files successfully deleted
	;  returns 0 & sets @error to 1 if path was invalid
	;  returns 0 if no files could be deleted
	$search = FileFindFirstFile($dir & "\*")
	if @error Then
		SetError(1)
		Return
	EndIf

	if $topLevel Then
		$temp = DirGetSize($dir, 1)
		$totalFiles = $temp[1]
		if $progress Then ProgressOn( "_delTree", $dir, "", 5, 5, 16 )
	EndIf

	While 1
		$current = FileFindNextFile($search)
		; exit the loop when there are no more files
		if @error Then ExitLoop
		$attrib = FileGetAttrib($dir & "\" & $current)
		; if directory, recurse into it and start the process over
		if StringInStr($attrib, "D") Then
			_delTree($dir & "\" & $current, $files, $progress, 0, 0, $totalFiles)
			; delete the directory (not possible if directory isn't empty, generally the case with dirs containing read-only or otherwise in-use files)
			DirRemove( $dir & "\" & $current )
		; delete the file otherwise and increment the count if successful
		Else
			$success = FileDelete( $dir & "\" & $current )
			if $success Then
				$files += 1
				if $progress Then ProgressSet( ($files/$totalFiles)*100, "Current file: " & $current & @CRLF & "Files deleted: " & $files, $dir )
				;TrayTip("", $dir & "\" & $current, 1)
			EndIf
		EndIf
		;ConsoleWrite($current & " attributes: " & $attrib & @CRLF )
	WEnd

	; release search handle at the end of each iteration
	FileClose($search)
	; return # of files deleted to calling delTree or main function
	if $topLevel Then
		if $progress Then ProgressOff()
		if $delTop Then FileDelete( $dir )
	EndIf
EndFunc

Func _fileFindDesktopShortcuts($targetSearch)
	Dim $dirSearch = FileFindFirstFile("c:\Documents and Settings\*")
	Dim $results

	While 1
		$current = FileFindNextFile($dirSearch)
		if @error Then ExitLoop

		$attrib = FileGetAttrib("c:\Documents and Settings\" & $current)
		if StringInStr($attrib, "D") Then
			$userSearch = FileFindFirstFile("c:\Documents and Settings\" & $current & "\Desktop\*.lnk")
			While 1
				$lnkSearch = FileFindNextFile($userSearch)
				if @error Then ExitLoop

				$scArray = FileGetShortcut("c:\Documents and Settings\" & $current & "\Desktop\" & $lnkSearch)
				if StringInStr($scArray[0], $targetSearch) Then
					$results &= "c:\Documents and Settings\" & $current & "\Desktop\" & $lnkSearch & @CRLF
					ConsoleWrite( $lnkSearch & @CRLF )
				EndIf
			WEnd
		EndIf
	WEnd

	Return $results
EndFunc

Func _setBGClinical()
;~ 	'delete existing files
	if FileExists(@StartupCommonDir & "\bginfocommon.lnk") Then
		FileDelete(@StartupCommonDir & "\bginfocommon.lnk")
   EndIf

   if FileExists(@SystemDir & "\bginfocommon.vbs") Then
	   FileDelete(@SystemDir & "\bginfocommon.vbs")
   EndIf

	If FileExists("\\txhealth.org\landesk\Softrep\BgInfo\bginfoclinical.vbs") Then
	   FileCopy("\\txhealth.org\landesk\Softrep\BgInfo\bginfoclinical.vbs", @SystemDir)
   EndIf

	$iReturn = FileCreateShortcut(@SystemDir & "\bginfoclinical.vbs", _
		@StartupCommonDir & "\bginfoclinical.lnk", _
		@SystemDir, "", "", "\\txhealth.org\landesk\Softrep\BgInfo\Window.ico", "", 0)

	Return $iReturn
EndFunc

Func _setBGCommon()
;~ 	'delete existing files
	if FileExists(@StartupCommonDir & "\bginfoclinical.lnk") Then
		FileDelete(@StartupCommonDir & "\bginfoclinical.lnk")
   EndIf

   if FileExists(@SystemDir & "\bginfoclinical.vbs") Then
	   FileDelete(@SystemDir & "\bginfoclinical.vbs")
   EndIf

	If FileExists("\\txhealth.org\landesk\Softrep\BgInfo\bginfocommon.vbs") Then
	   FileCopy("\\txhealth.org\landesk\Softrep\BgInfo\bginfocommon.vbs", @SystemDir)
   EndIf

	$iReturn = FileCreateShortcut(@SystemDir & "\bginfocommon.vbs", _
		@StartupCommonDir & "\bginfocommon.lnk", _
		@SystemDir, "", "", "\\txhealth.org\landesk\Softrep\BgInfo\Window.ico", "", 0)

	Return $iReturn
EndFunc

Func _getSCInfo($sSerial, $sField, $iVisible = 0)
	$sSCSearchURL = "http://serviceconnect/sm7/cwc/nav.menu?name=navStart&id=ROOT%2FMenu%20Navigation%2FConfiguration%20Management%2FResources%2FSearch%20CIs"

	$oIE = ObjCreate("InternetExplorer.Application")
	if NOT IsObj($oIE) Then Return SetError(3)

	$oIE.Visible = $iVisible
	_IENavigate($oIE, $sSCSearchURL)

	$oFrame = _IEFrameGetObjByName($oIE, "detail")
	$oForm = _IEFormGetObjByName($oFrame, "topaz")
	$oInputCI = _IEFormElementGetObjByName($oForm, "instance/logical.name")
	_IEFormElementSetValue($oInputCI, $sSerial)
	_IELinkClickByText($oFrame, " Search")
	Sleep(500)
	if WinExists("Windows Internet Explorer", "No records") Then
		WinClose("Windows Internet Explorer")
		$oIE.Quit
		Return SetError(1)
	Else
		$aFields = StringSplit($sField, "|")
		Dim $aReturn[$aFields[0]+1]

		$oForm = _IEFormGetObjByName($oFrame, "topaz")
		if _IEFormElementGetObjByName($oForm, " Count") Then
			$oIE.Visible = 1
			Return SetError(2)
		EndIf
		$oInputSerial = _IEFormElementGetObjByName($oForm, "instance/logical.name")
		$aReturn[0] = _IEFormElementGetValue($oInputSerial)

		For $i=1 To $aFields[0]
			$oElement = _IEFormElementGetObjByName($oForm, "instance/" & $aFields[$i])
			$aReturn[$i] = _IEFormElementGetValue($oElement)
		Next

		_IELinkClickByText($oFrame, " Cancel")
		$oIE.Quit
		Return $aReturn
	EndIf
EndFunc

Func _dirSize($dir, $filter = "", $updateFunc = False )
	; Returns # of files successfully deleted
	;  returns 0 & sets @error to 1 if path was invalid
	;  returns 0 if no files could be deleted
	Local $temp[3], $arrReturn[3]

	$search = FileFindFirstFile($dir & "\*")
	if @error Then
		SetError(1)
		Return
	EndIf

	While 1
		$current = FileFindNextFile($search)
		; exit the loop when there are no more files
		if @error Then ExitLoop
		$attrib = FileGetAttrib($dir & "\" & $current)
		; if directory, recurse into it and start the process over
		if StringInStr($attrib, "D") Then
			$temp = _dirSize($dir & "\" & $current, $filter)
			; add the total filesize to teh count
			$arrReturn[2] += 1
			if NOT IsArray($temp) Then ContinueLoop
			$arrReturn[0] += $temp[0]
			$arrReturn[1] += $temp[1]
			$arrReturn[2] += $temp[2]
		Else
			if $filter Then
				if NOT StringInStr( $current, $filter ) Then
					ContinueLoop
				Else
					ConsoleWrite($dir & "\" & $current & @CRLF)
				EndIf
			EndIf

			$arrReturn[1] += 1
			$arrReturn[0] += FileGetSize($dir & "\" & $current)
			if $updateFunc Then
				TrayTip($dir, "Total size: " & $arrReturn[0] & " bytes" & @CRLF & "Files: " & $arrReturn[1] & "Dirs: " & $arrReturn[2], 1)
				;GUICtrlSetData($labelSize, $arrReturn[1] & " files equalling " & Round($arrReturn[0]/1048576, 2) & " MB")
			EndIf
		EndIf
	WEnd

	; release search handle at the end of each iteration
	FileClose($search)
	; return values
	Return $arrReturn
EndFunc

Func guiFlash(ByRef $control, $color, $duration = 200, $times = 2, $tween = 0.4)
	$sleep1 = ($duration / $times) * $tween
	$sleep2 = ($duration / $times) * (1 - $tween)

	If $control <> "" Then
		If IsHWnd($control) Then
			For $i = 1 To $times
				GUISetBkColor($control, $color)
				Sleep($sleep1)
				GUISetBkColor($control, Default)
				Sleep($sleep2)
			Next
		Else
			For $i = 1 To $times
				GUICtrlSetBkColor($control, $color)
				Sleep($sleep1)
				GUICtrlSetBkColor($control, Default)
				Sleep($sleep2)
			Next
		EndIf
	Else
		Return 0
	EndIf

	Return 1
EndFunc   ;==>guiFlash

func timeFormat($ms, $bMS = False, $sepChar = ":")
	dim $sec, $min, $hr
	dim $s = round($ms/1000, 0 )

	$sec = mod($s, 60)
	if $sec < 10 then $sec = "0" & $sec
	$min = Floor(mod($s/60, 60))
	if $min < 10 then $min = "0" & $min
	$hr = Floor($s/3600)
	if $hr < 10 then $hr = "0" & $hr

	if $bMS Then
		$msTemp = Round(Mod($ms, 1000))
		$msTemp &= _StringRepeat("0", 3-StringLen($msTemp))
		Return $hr & $sepChar & $min & $sepChar & $sec & "." & $msTemp
	Else
		Return $hr & $sepChar & $min & $sepChar & $sec
	EndIf
EndFunc ; <== timeFormat

Func numberPadZeroesFloat($iNum, $iPlaces=2)
	if NOT IsNumber($iNum) OR NOT IsNumber($iPlaces) Then
		Return SetError(1)
	ElseIf IsInt($iNum) Then
		Return String($iNum & "." & _StringRepeat("0", $iPlaces))
	EndIf

	$aNum = StringSplit($iNum, ".", 2)

	For $i=1 To $iPlaces-StringLen($aNum[1])
		$iNum &= "0"
	Next

	Return String($iNum)
EndFunc

Func numberPadZeroesInt($iNum, $iPlaces=2)
	if NOT IsNumber($iNum) OR NOT IsNumber($iPlaces) Then
		Return SetError(1)
	ElseIf IsInt($iNum) Then
		$iLen = StringLen($iNum)
		if $iPlaces <= $iLen Then
			Return $iNum
		Else
			Return String(_StringRepeat("0", $iLen-$iPlaces) & $iNum)
		EndIf
	Else
		Return SetError(1)
	EndIf
EndFunc

Func _StringInsertRepeat($s, $sIns = ",", $iSpace = 3)
	Local $sNew
	For $i=Floor(StringLen($s)/$iSpace) To 3 Step -1
		$sNew &= StringMid($s, $i, 3) & $sIns
	Next
	Return $sNew
EndFunc

Func FileGetSizeAuto($sFile)
	if NOT FileExists($sFile) Then
		Return SetError(1)
	EndIf

	if StringInStr(FileGetAttrib($sFile), "D") Then
		$iSize = DirGetSize($sFile)
	Else
		$iSize = FileGetSize($sFile)
	EndIf

	Switch $iSize
		Case 1099511627775 to 1125899906842623
			Return numberPadZeroesFloat(Round($iSize/1073741824, 2)) & " TB"
		Case 1073741824 To 1099511627774
			Return numberPadZeroesFloat(Round($iSize/1073741824, 2)) & " GB"
		Case 1048576 To 1073741823
			Return numberPadZeroesFloat(Round($iSize/1048576, 2)) & " MB"
		Case 1024 To 1048575
			Return numberPadZeroesFloat(Round($iSize/1024, 2)) & " KB"
		Case Else
			Return $iSize & " B"
	EndSwitch
EndFunc

#cs
Func _guictrlbutton_createWithIcon($BItext, $BIleft, $BItop, $BIwidth, $BIheight, $sIcon, $BIconNum = 0)
    GUICtrlCreateIcon($sIcon, $BIconNum, $BIleft + 5, $BItop + (($BIheight - 16) / 2), 16, 16)
    GUICtrlSetState( -1, $GUI_DISABLE)
    $XS_btnx = GUICtrlCreateButton($BItext, $BIleft, $BItop, $BIwidth, $BIheight, $WS_CLIPSIBLINGS)
    Return $XS_btnx
EndFunc
#ce

Func _explorerNavSound($b = False)
	Switch $b
		Case True
			$iReturn = RegWrite("HKCU\AppEvents\Schemes\Apps\Explorer\Navigating\.current", "", "REG_EXPAND_SZ", "%SystemRoot%\media\Windows Navigation Start.wav")
		Case False
			$iReturn = RegWrite("HKCU\AppEvents\Schemes\Apps\Explorer\Navigating\.current", "", "REG_EXPAND_SZ", "")
	EndSwitch

	Return $iReturn
EndFunc

Func _fileGetIcon($sFile)
	; returns [iconPath, iconIndex, mimeType]
	$aType = StringSplit($sFile, ".")

	$sExt = $aType[$aType[0]]

	$sMIME = RegRead("HKCR\." & $sExt, "")
	if @error Then Return SetError(1)

	$sIcon = RegRead("HKCR\" & $sMIME & "\DefaultIcon", "")
	if @error Then Return SetError(2)

	$sIcon = StringReplace($sIcon, "%SystemRoot%", @WindowsDir)

	$iSplit = StringInStr($sIcon, ",", 0, -1)
	Local $aReturn[3] = [StringLeft($sIcon, $iSplit-1), StringTrimLeft($sIcon, $iSplit), $sMIME]

	Return $aReturn
EndFunc

Func _GUICreateSimpleConsole(ByRef $hWnd, ByRef $hEdit, $sTitle = "Console", $iX = 669, $iY = 333, $style = 0x00040000, $exstyle = -1, $hParent = -1)
	$hWnd = GUICreate( $sTitle, $iX, $iY, @DesktopWidth/2, @DesktopHeight/2, $style, $exstyle, $hParent)
	$hEdit = GUICtrlCreateEdit( "", 0, 0, $iX, $iY, BitOR(4,2048,0x00200000) )
	GUICtrlSetCursor(-1, 2)
	GUICtrlSetBkColor( -1, 0x000000 )
	GUICtrlSetFont( -1, 8, 400, 0, "Fixedsys" )
	GUICtrlSetColor( -1, 0x888888)
	GUICtrlSetResizing( -1, 0x0066)

	Return 1
EndFunc