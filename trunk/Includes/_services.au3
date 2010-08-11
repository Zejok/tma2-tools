; #INDEX# =======================================================================================================================
; Title .........: _services
; AutoIt Version : 3.3.6.1
; Description ...: UDF for misc. computer management tasks
; Company .......: Mosaic Technology
; Author ........: Jon Dunham
; To-do .........: MUST ADD THE REFCOUNT FOR ANY INCLUDING SCRIPT:
;                  Global $oRefCount = ObjCreate("WbemScripting.SWbemNamedValueSet")
;                  add(), remove(), item(), deleteall() & count
; ===============================================================================================================================
#include <file.au3>
#include <array.au3>
#include-once

#cs
$oWbemSink = ObjCreate("WbemScripting.SWbemSink")
ObjEvent($oWbemSink, "__WbemSink_")

Global $oRefCount = ObjCreate("WbemScripting.SWbemNamedValueSet")

$hTotalTime = TimerInit()

Local $aSites
_FileReadToArray(@DesktopDir & "\testHosts.txt", $aSites)
#ce

#cs
Local $aStats
$cpuStats1 = ProcessGetStats(-1, 0)
_ArrayInsert($cpuStats1, 0, " -----[ Process Memory")
$cpuStats2 = ProcessGetStats(-1, 1)
_ArrayInsert($cpuStats2, 0, " -----[ Process CPU Usage")
$memStats = MemGetStats()
_ArrayInsert($memStats, 0, " -----[ Global Memory")

_ArrayConcatenate($aStats, $cpuStats1)
_ArrayConcatenate($aStats, $cpuStats2)
_ArrayConcatenate($aStats, $memStats)
_ArrayDisplay($aStats, "Before")


For $line = 1 to $aSites[0]
	$oPing = _superPing($aSites[$line], $line, $oWbemSink, 100)
	$oRefCount.Add($aSites[$line], $oPing)

	Sleep(500)
Next

$memStats = MemGetStats()
ConsoleWrite("-----[ Global Memory Use: " & $memStats[0] & @CRLF)

While $oRefCount.Count > 0
	For $oSiteRef in $oRefCount
		$oShit = $oSiteRef.Value
		ConsoleWrite($oShit.Item("Address").Value & " (" & $oShit.Item("ResolvedIP").Value & ") " & @TAB & " status: " & $oShit.Item("Status").Value & ". Reply: " & $oShit.Item("Reply").Value & @CRLF)
		$oShit.DeleteAll()
	Next
WEnd

$oRefCount.DeleteAll()

ConsoleWrite("+Total ping time: " & Round(TimerDiff($hTotalTime)) & @CRLF)
#ce

Func _superPing($sAddr, ByRef $oGlobalSink, $iTimeout = 200, $lParam = -1)
	; ========== NOTE!!!
	; when everything starts becoming asynchronous, be sure to monitor # of calls being made, instances of wmiprvse and resultant CPU usage
	; YES, YES, A THOUSAND TIMES YES!!
	;
;~ 	Local $wbemFlagReturnImmediately = 16, $wbemFlagForwardOnly = 32
	Local $wbemFlagSendStatus = 128
	$oWMILocal = ObjGet("winmgmts:\\.\root\cimv2")

	$oNamedValueSet = ObjCreate("WbemScripting.SWbemNamedValueSet")
	$oNamedValueSet.Add("Address", $sAddr)
	$oNamedValueSet.Add("PingStatus", -1)
;~ 	$oNamedValueSet.Add("PercentDone", 0) ; use for future functions that return > 1 object
	$oNamedValueSet.Add("Caller", "PingStatus")
	$oNamedValueSet.Add("lParam", $lParam)

	$oWMILocal.ExecQueryAsync($oGlobalSink, 'SELECT * FROM Win32_PingStatus WHERE Address="' & $sAddr & '" AND Timeout=' & $iTimeout & ' AND ResolveAddressNames=True', "WQL", $wbemFlagSendStatus, Default, $oNamedValueSet)

	Return $oNamedValueSet
EndFunc

#cs
Func __WbemSink_OnProgress($iUpperBound, $iCurrent, $sMessage, $oAsyncContext)
	if $iUpperBound <= 1 Then Return

	Switch $oAsyncContext.Item("Caller").Value
		Case "PingStatus"
			ConsoleWrite("+Progress: " & $iCurrent & "/" & $iUpperBound & " (" & ($iCurrent/$iUpperBound)*100 & "%)" & @CRLF)
			ConsoleWrite("+Message: " & $sMessage & @CRLF)
		Case "Test"
			ConsoleWrite("+Progress: " & $iCurrent & "/" & $iUpperBound & " (" & ($iCurrent/$iUpperBound)*100 & "%)" & @CRLF)
			ConsoleWrite("+Message: " & $sMessage & @CRLF)
	EndSwitch
EndFunc
#ce
#cs
Func __WbemSink_OnObjectReady($oObject, $oAsyncContext)
	; use "Caller" for all script calls
	if $oAsyncContext.Item("Caller").Value = "PingStatus" Then
;~ 		$oPingProperties = $oObject.Properties_
		$oAsyncContext.Add("Reply", $oObject.ResponseTime)
		$oAsyncContext.Add("ResolutionStatus", $oObject.PrimaryAddressResolutionStatus)
		$oAsyncContext.Add("ResolvedAddress", $oObject.ProtocolAddress)
		$oAsyncContext.Add("ResolvedIP", $oObject.ProtocolAddressResolved)

		$oAsyncContext.Item("Status").Value = $oObject.StatusCode
	ElseIf $oAsyncContext.Item("Caller").Value = "Test" Then
		ConsoleWrite("!Test object ready" & @CRLF)
	EndIf

	if IsDeclared("oRefCount") Then
;~ 		$oRefCount.Item("RefCount").Value = $oRefCount.Item("RefCount").Value - 1
	EndIf
EndFunc
#ce

Func _regEnumKeys($sKey, $sComp = ".")
	const $HKEY_LOCAL_MACHINE = 0x80000002
	Local $arrSubKeys
	Local $objReg=ObjGet("winmgmts:{impersonationLevel=impersonate}!\\"& $sComp & "\root\default:StdRegProv")
	if @error Then Return SetError(1)

	$objReg.EnumKey($HKEY_LOCAL_MACHINE, $sKey, $arrSubKeys)
	Return $arrSubKeys
EndFunc

Func _checkCert($sComp = ".")
	Local $sKey = "\\" & $sComp & "\HKLM\SOFTWARE\Microsoft\SystemCertificates\My\Certificates"
	Local $i = 1
	Local $aProperties[5]
	Local $oCert = ObjCreate("CAPICOM.Certificate.2")
	if @error Then
		Return SetError(4)
	EndIf

	While 1
		$sEnum = RegEnumKey($sKey, $i)
		Switch @error
			Case -1
				; no certificates or none matched the fingerprint offset
				if $i = 1 Then
					Return SetError(1)
				Else
					Return SetError(2)
				EndIf
			Case 1 to 2
				; probably rare
				Return SetError(4, @error)
			Case 3
				Return SetError(3)
			Case 0
				$binaryBlob = RegRead($sKey & "\" & $sEnum, "blob")
				; needs to be REG_BINARY
				if @extended <> 3 Then
					$i += 1
					ContinueLoop
				EndIf

				$hKey = Binary("0x" & $sEnum)
				$hMid = BinaryMid($binaryBlob, Int(0x0180)+1, 20)

				if $hMid = $hKey Then
					$oCert.Import($binaryBlob)
					$err = @error ; assumes calling script handles & sets COM errors

					if $err Then
						$oCert = 0
						Return SetError(5, $err)
					EndIf

					$oCertStatus = $oCert.IsValid()
					$aProperties[0] = $oCert.SubjectName
					$aProperties[1] = $oCert.IssuerName
					$aProperties[2] = $oCert.ValidFromDate
					$aProperties[3] = $oCert.ValidToDate
					$aProperties[4] = $oCertStatus.Result

					$oCert = 0
					$oCertStatus = 0
					Return $aProperties
				EndIf
		EndSwitch
		$i += 1
	WEnd
EndFunc

Func _AppvGetApps($sComp = ".") ; name, version, last launch, pkg GUID, source OSD, global running count
	$oWMI = ObjGet("winmgmts:\\" & $sComp & "\root\microsoft\appvirt\client")
	if NOT IsObj($oWMI) Then
		Return SetError(1)
	EndIf

	; return 0 if no application instances
	$colApps = $oWMI.ExecQuery("Select * from Application")
	if $colApps.Count = 0 Then Return 0

	Dim $objDateTime = ObjCreate("WbemScripting.SWbemDateTime")
	if NOT IsObj($objDateTime) Then
		Return SetError(2)
	EndIf

	Dim $aReturn[$colApps.Count+1][6]
	Dim $i=1

	$aReturn[0][0] = $colApps.Count

	For $objApp in $colApps
		$objDateTime.Value = $objApp.LastLaunchOnSystem

		$aReturn[$i][0] = $objApp.Name
		$aReturn[$i][1] = $objApp.Version

;~ 		ConsoleWrite("datetime " & $objDateTime.Value & @CRLF)
		$aReturn[$i][2] = $objDateTime.Month & "/" & $objDateTime.Day & "/" & $objDateTime.Year & " " & $objDateTime.Hours & ":" & $objDateTime.Minutes

		$aReturn[$i][3] = $objApp.PackageGUID
		$aReturn[$i][4] = $objApp.OriginalOSDPath
		$aReturn[$i][5] = $objApp.GlobalRunningCount
		$i += 1
	Next

	Return $aReturn
EndFunc

Func _renameComp($SNEWNAME, $SCOMP = @ComputerName, $USERNAME = Default, $PASS = Default)
	$OWMI = ObjGet("winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!\\" & $SCOMP & "\root\cimv2")
	$COLCS = $OWMI.ExecQuery("Select * from Win32_ComputerSystem")
	For $OBJCS In $COLCS
		$IRETURN = $OBJCS.Rename($SNEWNAME, $PASS, $USERNAME)
	Next
	$OWMI = 0
	If $IRETURN Then
		Return SetError($IRETURN)
	Else
		Return 1
	EndIf
EndFunc

Func _serviceControl($sServiceName, $sAction, $sArg = "", $sComp= ".")
	; ChangeStartMode, StartService, StopService etc
	Dim $oWMI = ObjGet("winmgmts:\\" & $sComp & "\root\cimv2"), $iReturn = 1, $bExists = False

	$col = $oWMI.ExecQuery('SELECT * FROM Win32_Service WHERE Name="' & $sServiceName & '"')

	for $obj in $col
		$bExists = True

		Switch $sAction
		Case "ChangeStartMode"
			$iReturn = $obj.ChangeStartMode($sArg)
		Case "StartService"
			$iReturn = $obj.StartService()
		Case "StopService"
			$iReturn = $obj.StopService()
		Case "InterrogateService"
			$iReturn = $obj.InterrogateService()
		Case Else
			Return SetError(2)
		EndSwitch
	Next
	$oWMI = 0

	if NOT $bExists Then Return SetError(1)

	if NOT $iReturn Then
		Return 1
	Else
		Return SetError(3, $iReturn)
	EndIf
EndFunc

Func _processStart($sCommand, $sComp = ".", $iShowWindow = 0)
	Dim $iPID = 0

	$oWMI = ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & $sComp & "\root\cimv2")

	$oStartup = $oWMI.Get("Win32_ProcessStartup")
	$oConfig = $oStartup.SpawnInstance_
	$oConfig.ShowWindow = $iShowWindow

	$oProcess = $oWMI.Get("Win32_Process")

	$iReturn = $oProcess.Create($sCommand, Default, $oConfig, $iPID)

	if $iReturn <> 0 Then
		Return SetError($iReturn)
	Else
		Return $iPID
	EndIf

EndFunc

Func CPUArch($sComp = ".")
	$oWMI = ObjGet("winmgmts:\\" & $sComp & "\root\cimv2")
	if NOT IsObj($oWMI) Then
		Return SetError(1)
	EndIf

	$oWMI.Get('Win32_Processor.DeviceID="CPU0"')

	Switch $oWMI.Architecture
		Case 0
			Return "X86"
		Case 6 to 9
			Return "X64"
		Case Else
			Return $oWMI.Architecture
	EndSwitch

EndFunc

Func _processExists($procName, $compname = ".")
	$oWMIService = ObjGet( "winmgmts:\\" & $compName & "\root\CIMV2" )

	if NOT IsObj($oWMIService) Then
		SetError(1)
		Return
	EndIf

	Dim $handle = False, $cProc

	Switch VarGetType($procName)
	Case "String"
		$cProc = $oWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE Name = "' & $procName & '"')
	Case "Int32"
		$cProc = $oWMIService.ExecQuery('SELECT * FROM Win32_Process WHERE ProcessId = "' & $procName & '"')
	Case Else
		Return SetError(2)
	EndSwitch

	for $oProc in $cProc
		$handle = $oProc.Handle
	Next

	if $handle <> False Then
		Return $handle
	Else
		Return 0
	EndIf
EndFunc

Func _netSetStaticIP($sIP, $sSubnet, $sGateway, $sComp = ".")
	;Win32_NetworkAdapterConfiguration

	Dim $iReturn = -1
	Dim $oWMI = ObjGet("winmgmts:\\" & $sComp & "\root\cimv2")

	$col = $oWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")

	Dim $arIP[1] = [$sIP]
	Dim $arSubnet[1] = [$sSubnet]
	Dim $arGateway[1] = [$sGateway]
	Dim $arGatewayMetric[1] = [1]

	for $obj in $col
		$obj.EnableStatic($arIP, $arSubnet)
		$obj.SetGateways($arGateway, $arGatewayMetric)
	Next

	$oWMI = 0
EndFunc

Func _netGetDNS($sComp = ".")
	Dim $aReturn
	Dim $oWMI = ObjGet("winmgmts:\\" & $sComp & "\root\cimv2")

	$col = $oWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")

	For $obj in $col
		$aReturn = $obj.DNSServerSearchOrder
	Next

	$oWMI = 0

	if IsArray($aReturn) Then
		Return $aReturn
	Else
		Return SetError(1)
	EndIf
EndFunc

Func _netGetWINS($sComp = ".")
	Dim $aReturn[2]
	Dim $oWMI = ObjGet("winmgmts:\\" & $sComp & "\root\cimv2")

	$col = $oWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")

	For $obj in $col
		$aReturn[0] = $obj.WINSPrimaryServer
		$aReturn[1] = $obj.WINSSecondaryServer
	Next

	$oWMI = 0

	if $aReturn[0] <> "" Then
		Return $aReturn
	Else
		Return SetError(1)
	EndIf
EndFunc

Func _netSetDNS($arDNS, $sComp = ".")
	Dim $oWMI = ObjGet("winmgmts:\\" & $sComp & "\root\cimv2")

	$col = $oWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")

	For $obj in $col
		$iReturn = $obj.SetDNSServerSearchOrder($arDNS)
	Next

	Return $iReturn
EndFunc

Func _netSetWINS($sWINS1, $sWINS2, $sComp = ".")
	Dim $iReturn
	Dim $oWMI = ObjGet("winmgmts:\\" & $sComp & "\root\cimv2")

	$col = $oWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")

	For $obj in $col
		$iReturn = $obj.SetWINSServer($sWINS1, $sWINS2)
	Next

	$oWMI = 0

	Return $iReturn
EndFunc

Func _netValidIP($strIP)
	Return StringRegExp($strIP, "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b")
EndFunc