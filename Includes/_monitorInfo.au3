#include <_services.au3>
#include <file.au3>
#include-once

Func _monitorInfoSemiSync($compName = ".")
	; return colMonitors, with items (DeviceID) set to
	;   objMonitor, with properties:
	;   Availability, DeviceID, Name, PNPDeviceID, Status, EDIDBinary

	$wbemFlagReturnImmediately = 16
	$wbemFlagForwardOnly = 32

	Local $objWMIService = ObjGet("winmgmts:\\" & $compName & "\root\cimv2")
	if @error Then
		Return SetError(1)
	EndIf

	$colMonitors = ObjCreate("Wbemscripting.SWbemNamedValueSet")
	$colSysDM = $objWMIService.ExecQuery("Select Availability, DeviceID, Name, PNPDeviceID, Status " & _
		"From Win32_DesktopMonitor " & _
		"WHERE PNPDeviceID IS NOT NULL AND (Availability = 3 OR Availability = 8)", "WQL", $wbemFlagForwardOnly+$wbemFlagReturnImmediately)

	for $oSysDM in $colSysDM
		$tempEDID = RegRead("\\" & $compName & "\HKLM\SYSTEM\CurrentControlSet\Enum\" & $oSysDM.PNPDeviceID & "\Device Parameters", "EDID")
		if @error Then ContinueLoop

		$objMonitor = ObjCreate("Wbemscripting.SWbemNamedValueSet")
		$objMonitor.Add("DeviceID", $oSysDM.DeviceID)
		$objMonitor.Add("Availability", $oSysDM.Availability)
		$objMonitor.Add("Name", $oSysDM.Name)
		$objMonitor.Add("PNPDeviceID", $oSysDM.PNPDeviceID)
		$objMonitor.Add("Status", $oSysDM.Status)
		$objMonitor.Add("EDIDBinary", $tempEDID)

		$colMonitors.Add($objMonitor.DeviceID, $oSysDM)
	Next

	Return $colMonitors
EndFunc

Func _monitorInfoFromEDID($EDID)
	Global Const $EDID_HEADER = Binary("0x00FFFFFFFFFFFF00")




EndFunc


Func _monitorInfo($compName = ".", $iStyle = 1)
	; return colMonitors, with items (DeviceID) set to
	;   objMonitor, with properties:
	;   Availability, DeviceID, Name, PNPDeviceID, Status, EDIDBinary
	$colMonitors = ObjCreate("Wbemscripting.SWbemNamedValueSet")

	$sPathEDID = '"\\ftwgen01\THIS\Field Services\autoItScripts\psTools\dumpedid.exe"'
	if $iStyle = 2 Then
		if NOT FileExists($sPathEDID) Then
			$sPathEDID = @TempDir & "\dumpedid.exe"
			FileInstall("dumpedid.exe", $sPathEDID)
		EndIf
	EndIf
	; Dell serial - 78 characters in (without MX0 or CN0 prefix)
	; HP serial - 114 characters in
	; both models - 96 (192 hex) characters in
	;
	; this function runs assuming the computer has already been contacted, otherwise it will hang for ~82 seconds trying to contact the WMI service
	;
	Dim $colDM, $PNPDID, $EDID, $Name, $iAvail = 0
	Dim $objWMIService = ObjGet("winmgmts:\\" & $compName & "\root\cimv2")
	if NOT IsObj($objWMIService) Then
		SetError(1)
		Return
	EndIf

	$colDM = $objWMIService.ExecQuery('SELECT * FROM Win32_DesktopMonitor WHERE PNPDeviceID IS NOT NULL')

	for $objDM in $colDM
		if $objDM.Availability = 3 Then ; this is the best scenario we want (powered on/connected). Generally all other DesktopMonitor.Availability = 8 (off-line)
			$PNPDID = $objDM.PNPDeviceID
			$Name = $objDM.Name
			$EDID = RegRead("\\" & $compName & "\HKLM\SYSTEM\CurrentControlSet\Enum\" & $PNPDID & "\Device Parameters", "EDID" )
			if NOT @error Then
				$iAvail = $objDM.Availability
				ExitLoop
			Else
				ContinueLoop ; if there isn't an EDID for this, continueloop and get it from the next monitor with a PNPDeviceID
			EndIf
		Else ; this may or may not indicate the currently connected display device. Display status reporting seems to be sketchy at best with WMI
			$PNPDID = $objDM.PNPDeviceID
			$Name = $objDM.Name
			$EDID = RegRead("\\" & $compName & "\HKLM\SYSTEM\CurrentControlSet\Enum\" & $PNPDID & "\Device Parameters", "EDID" )
			if NOT @error Then
				$iAvail = $objDM.Availability
				ExitLoop
			Else
				ContinueLoop ; if there isn't an EDID for this, continueloop and get it from the next monitor with a PNPDeviceID
			EndIf
		EndIf
	Next

	; sample EDID
	; DISPLAY\ELO0886\4&313E5566&0&80861500&00&02
;~ 	if $debug Then ConsoleWrite( $Name & ": " & $objDM.Availability & " \ " & $PNPDID & @CRLF )

	Switch $iStyle
	Case 1
		if NOT IsBinary($EDID) Then
			SetError(1)
			Return @error
		EndIf

		Dim $info[4], $serial, $model
		$info[3] = $iAvail

		$model = StringStripWS(StringMid(BinaryToString($EDID), 96, 12), 2)

	;~ 	if $debug then ConsoleWrite( "Model: " & $model & @CRLF )

		Select
		Case StringLeft($model, 2) = "HP"
	;~ 		if $debug then ConsoleWrite( "HP found" & @CRLF )
			$serial = StringMid(BinaryToString($EDID), 114, 10)
		Case StringLeft($model, 2) = "DE"
	;~ 		if $debug then ConsoleWrite( "Dell found" & @CRLF )
			$serial = StringMid(BinaryToString($EDID), 78, 12)
			$serial = StringLeft($serial, 5) & " | " & StringRight($serial, 7)
		Case StringLeft($model, 2) = "LG"
	;~ 		if $debug then ConsoleWrite( "LG found" & @CRLF )
			$serial = StringMid(BinaryToString($EDID), 114, 12) & " (may be the model)"
		EndSelect

	;~ 	if $debug then ConsoleWrite( "Serial: " & $serial & @CRLF )

		$info[0] = $model
		$info[1] = $serial
		$info[2] = $Name
	Case 2
		; 0 / 1 / 2 / 3 / 4 / 5 / 6
		; model / serial / PNP name / full dump / # serial / availability / maximum image size
		Dim $info[7]
		$info[5] = $iAvail
		$sFilePath = @TempDir & "\" & $compName & "-EDID.txt"
		$iReturn = ShellExecuteWait(@ComSpec, '/c ' & $sPathEDID & ' ' & $compName & ' > "' & $sFilePath & '"', "", "", @SW_HIDE)
		$iLength = _FileCountLines($sFilePath)*2
		$hFile = FileOpen($sFilePath, 0)

		for $iStart=1 To $iLength Step 2
			$sLine = FileReadLine($hFile, $iStart)
			$iString = StringInStr($sLine, $PNPDID, 2)
			if $iString Then
				For $iEnd=$iStart To $iLength Step 2
					$iString = StringInStr(FileReadLine($hFile, $iEnd), "**********")
					if $iString Then
						$iRecordLength = ($iEnd-1)-($iStart-1)
						ExitLoop
					EndIf
				Next

				For $i=$iStart To $iEnd Step 2
					$sLine = FileReadLine($hFile, $i)
					if NOT StringInStr($sLine, "**********") AND $sLine <> "" Then $info[3] &= $sLine & @CRLF
					if StringInStr($sLine, "Monitor Name") Then
						;model
						$info[0] = StringTrimLeft(FileReadLine($hFile, $i), 27)
					ElseIf StringInStr($sLine, "Serial Number   ") Then
						;serial
						$info[1] = StringTrimLeft(FileReadLine($hFile, $i), 27)
					ElseIf StringInStr($sLine, "Serial Number (Numeric)") Then
						;num. serial
						$info[4] = StringTrimLeft(FileReadLine($hFile, $i), 27)
					ElseIf StringInStr($sLine, "Maximum Image Size") Then
						;screen size
						$info[6] = Number(StringTrimRight(StringTrimLeft(FileReadLine($hFile, $i), StringInStr(FileReadLine($hFile, $i), "(")), 6))
					EndIf
				Next
			EndIf
		Next

		if StringInStr($info[0], "DELL") AND $info[1] <> "" Then _
			$info[1] = StringLeft($info[1], 5) & " | " & StringRight($info[1], 7)
		$info[2] = $name

		FileClose($hFile)
		FileDelete($sFilePath)
		FileDelete(@TempDir & "\dumpedid.exe")
	EndSwitch

	Return $info
EndFunc