#include-once
;~ #include <defPrinterInfo.au3>
#include <tma2.au3>


Func queryComp($compName = @ComputerName)
	; init object variables
	dim $objWMIService, $objRegistry
	dim $colBios, $colCSP, $colLMC, $colSysEnc

	; init arrays we'll return
	; $info[0] = model
	; $info[1] = serial
	; $info[2] = asset
	; $info[3] = bios
	; $info[4] = RAM
	; $info[5] = def user
	; $info[6] = OU
	; $info[7] = user
	; $info[8] = def. printer for user
	; $info[9] = MAC address
	; $info[10] = make

	dim $info[11], $SID

	; get the WMI object
	$objWMIService = ObjGet("winmgmts:\\" & $compName & "\root\cimv2")
	$info[5] = RegRead( "\\" & $compName & "\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinLogon", "DefaultUserName" )

	; check to see if the WMI object exists; if not, seterror and return
	; this check should now be deprecated due to the use of ping()
	if NOT IsObj( $objWMIService ) Then
		SetError(9)
		Return
	EndIf

	; execquery on all the info groups we'll need
	; BIOS for Serial and BIOS ver
	$colBios = $objWMIService.execquery("Select * From Win32_BIOS")
	; SysEnc for Asset
	$colSysEnc = $objWMIService.execquery("Select * From Win32_SystemEnclosure")
	; CompSystem for Model and current user
	$colCSP = $objWMIService.execquery("Select * from Win32_ComputerSystem")
	; LMC for total RAM
	$colLMC = $objWMIService.execquery("Select * from Win32_LogicalMemoryConfiguration")

	$cItems = $objWMIService.execquery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True" )
	for $oItem in $cItems
		$arrIP = $oItem.IPAddress
		if $arrIP[0] = TCPNameToIP($compName) Then
			$info[9] = $oItem.MACAddress
		EndIf
	Next

	$info[6] = _getDN($compName)

	; grab computer model
	for $objCSP in $colCSP
		$info[0] = $objCSP.Model
		$info[10] = $objCSP.Manufacturer
		$info[7] = $objCSP.UserName
	Next

	if $info[7] <> "" Then
		$arrCU = StringSplit($info[7], "\", 2)
		$objAccount = $objWMIService.Get('Win32_UserAccount.Name="' & $arrCU[1] & '",Domain="' & $arrCU[0] & '"')
;~ 		if IsObj($objAccount) Then
;~ 			$SID = $objAccount.SID
;~ 		EndIf
	EndIf

	$aPrinter = _defPrinterInfo($compName)
	if @error = 2 Then
		$info[8] = "No printers installed."
	ElseIf @error = 3 Then
		$info[8] = "No default set."
	ElseIf NOT @error Then
		$info[8] = $aPrinter[0]
	EndIf

	; grab the serial and BIOS version
	for $objBios in $colBios
		$info[1] = $objBios.SerialNumber
		$info[3] = $objBios.SMBIOSBIOSVersion
	Next

	; grab the asset tag, if available (it will be on HP dc7900s and possibly others)
	for $objSysEnc in $colSysEnc
		$info[2] = StringStripWS($objSysEnc.SMBIOSAssetTag, 8)
	Next

	; HP Asset if needed
	if StringInStr($info[0], "HP Compaq 8000 Elite") AND $info[2] = "" Then
		Dim $oWMIService2 = ObjGet("winmgmts:{impersonationlevel=impersonate}//" & $compName & "/root/HP/InstrumentedBIOS")
		if NOT IsObj($oWMIService2) Then
			$info[2] = "Blank or serial dupe"
		Else
			$temp = $oWMIService2.Get('HPBIOS_BIOSString.InstanceName="ACPI\\PNP0C14\\0_13"')
			$info[2] = StringStripWS($temp.Value, 8)

			if NOT $info[2] OR StringIsSpace($info[2]) Then $info[2] = "Blank or serial dupe"
			$temp = 0
		EndIf

		$oWMIService2 = 0
	ElseIf $info[2] = "" Then
		$info[2] = "Blank or serial dupe"
	EndIf

	; set to N/A if it didn't exist for returning purposes
	if StringIsSpace($info[2]) Then $info[2] = "Asset tag not set"

	; grab RAM, convert to GB ('cause who the fuck has less than 1GB nowadays)
	for $objLMC in $colLMC
		$info[4] = round($objLMC.TotalPhysicalMemory/1048576, 2)
	Next

	; Strip whitespace from end of model string, 'cause there's usually a lot
	$info[0] = StringStripWS($info[0], 2)

	Return $info

EndFunc

func compInfo($compName = @ComputerName)
	; init object variables
	dim $objWMIService, $colBios, $colCSP, $colSysEnc

	; init array we'll return and misc.
	; 0 manu 1 model 2 serial 3 asset 4 bios 5 ip
	dim $info[7], $DHCP, $index

	; get IP seen by the network
	TCPStartup()
	$info[5] = TCPNameToIP( $compName )
	TCPShutdown()

	; get the WMI object
	$objWMIService = ObjGet("winmgmts:\\" & $compName & "\root\cimv2")

	; execquery on all the info groups we'll need
	$colBios = $objWMIService.execquery("Select * From Win32_BIOS")
	$colSysEnc = $objWMIService.execquery("Select * From Win32_SystemEnclosure")
	$colCSP = $objWMIService.execquery("Select * from Win32_ComputerSystem")
	$colNAC = $objWMIService.execquery("Select * from Win32_NetworkAdapter WHERE NetConnectionStatus = 2")

	for $objNAC in $colNAC
		$index = $objNAC.Index
	Next

	$colNic = $objWMIService.execquery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE Index = " & $index )

	;DHCP
	for $objNic in $colNic
		; check if static IP or not, indicating QS or otherwise special-purpose machine
;~ 		$info[6] = $oItem.MACAddress
		$DHCP = $objNic.DHCPEnabled
		if $DHCP = 0 Then
			$info[5] &= " (Static)"
		Else
			$info[5] &= " (Dynamic)"
		EndIf
	Next

	; grab computer model
	for $objCSP in $colCSP
		$info[0] = StringStripWS($objCSP.Manufacturer, 2)
		$info[1] = StringStripWS($objCSP.Model, 2)
	Next

	; grab the serial and BIOS version
	for $objBios in $colBios
		$info[2] = $objBios.SerialNumber
		$info[4] = $objBios.SMBIOSBIOSVersion
	Next

	; grab the asset tag, if available (it will be on HP dc7900s and possibly others)
	for $objSysEnc in $colSysEnc
		$info[3] = StringStripWS($objSysEnc.SMBIOSAssetTag, 8)
	Next

	; HP Asset if needed
	if StringInStr($info[1], "HP Compaq 8000 Elite") AND $info[3] = "" OR $info[3] = $info[2] Then
		Dim $oWMIService2 = ObjGet("winmgmts:{impersonationlevel=impersonate}//" & $compName & "/root/HP/InstrumentedBIOS")
		if NOT IsObj($oWMIService2) Then
			$info[3] = "Blank or serial dupe"
		Else
			$temp = $oWMIService2.Get('HPBIOS_BIOSString.InstanceName="ACPI\\PNP0C14\\0_13"')
			$info[3] = StringStripWS($temp.Value, 8)

			if NOT $info[3] OR StringIsSpace($info[3]) Then $info[3] = "Blank or serial dupe"
			$temp = 0
		EndIf

		$oWMIService2 = 0
	ElseIf $info[3] = "" Then
		$info[3] = "Blank or serial dupe"
	EndIf

	$info[6] = RegRead("\\" & $compName & "\HKLM\SOFTWARE\LANDesk\ManagementSuite\WinClient\AMT", "GUID")

	; get wireless info if available
	#cs
	$infoW = wirelessInfo($compName)

	; ssid
	$info[12] = $infoW[0]
	; bssid
	$info[13] = $infoW[1]
	; signal
	$info[14] = $infoW[3]
	; channel
	$info[15] = $infoW[4]
	#ce

	Return $info

EndFunc

Func monitorInfo($compName = ".")
	; Dell serial - 78 characters in (without MX0 or CN0 prefix)
	; HP serial - 114 characters in
	; both models - 96 (192 hex) characters in
	;
	; this function runs assuming the computer has already been contacted, otherwise it will hang for ~82 seconds trying to contact the WMI service
	;
	Dim $colDM, $PNPDID, $EDID, $Name
	Dim $objWMIService = ObjGet("winmgmts:\\" & $compName & "\root\cimv2")
	if NOT IsObj($objWMIService) Then
		SetError(1)
		Return
	EndIf

	$colDM = $objWMIService.execquery('SELECT * FROM Win32_DesktopMonitor WHERE PNPDeviceID IS NOT NULL')

	for $objDM in $colDM
		if $objDM.Availability = 3 Then ; this is the best scenario we want (powered on/connected). Generally all other DesktopMonitor.Availability = 8 (off-line)
			$PNPDID = $objDM.PNPDeviceID
			$Name = $objDM.Name
			$EDID = RegRead("\\" & $compName & "\HKLM\SYSTEM\CurrentControlSet\Enum\" & $PNPDID & "\Device Parameters", "EDID" )
			if NOT @error Then
				ExitLoop
			Else
				ContinueLoop ; if there isn't an EDID for this, continueloop and get it from the next monitor with a PNPDeviceID
			EndIf
		Else ; this may or may not indicate the currently connected display device. Display status reporting seems to be sketchy at best with WMI
			$PNPDID = $objDM.PNPDeviceID
			$Name = $objDM.Name
			$EDID = RegRead("\\" & $compName & "\HKLM\SYSTEM\CurrentControlSet\Enum\" & $PNPDID & "\Device Parameters", "EDID" )
			if NOT @error Then
				ExitLoop
			Else
				ContinueLoop ; if there isn't an EDID for this, continueloop and get it from the next monitor with a PNPDeviceID
			EndIf
		EndIf
	Next

	;if $debug Then ConsoleWrite( $Name & ": " & $objDM.Availability & " \ " & $PNPDID & @LF & @LF & $EDID & @LF )

	if NOT IsBinary($EDID) Then
		SetError(1)
		Return @error
	EndIf

	Dim $info[3], $serial, $model

	$model = StringStripWS(StringMid(BinaryToString($EDID), 96, 12), 2)

	if $debug then ConsoleWrite( "Model: " & $model & @CRLF )

	Select
	Case StringLeft($model, 2) = "HP"
		if $debug then ConsoleWrite( "HP found" & @CRLF )
		$serial = StringMid(BinaryToString($EDID), 114, 10)
	Case StringLeft($model, 2) = "DE"
		if $debug then ConsoleWrite( "Dell found" & @CRLF )
		$serial = StringMid(BinaryToString($EDID), 78, 12)
		$serial = StringLeft($serial, 5) & "|" & StringRight($serial, 7)
	Case StringLeft($model, 2) = "LG"
		if $debug then ConsoleWrite( "LG found" & @CRLF )
		$serial = StringMid(BinaryToString($EDID), 114, 12) & " (may be the model)"
	EndSelect

	if $debug then ConsoleWrite( "Serial: " & $serial & @CRLF )

	$info[0] = $model
	$info[1] = $serial
	$info[2] = $Name

	Return $info
EndFunc