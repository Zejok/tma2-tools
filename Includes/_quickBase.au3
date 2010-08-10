#include-once

#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <buttonConstants.au3>

#include <inet.au3>
#include <ie.au3>
#include <date.au3>
#include <_XMLDomWrapper.au3>

#include <array.au3>

if NOT IsDeclared("debug") Then Global $debug = 1

#cs
PERSONAL QUICKBASE REFERENCE STUFF

#ce


;~ ====================================================
;~ -- THAM Healthcheck
;~ ====================================================

;~ Global Const $sAppToken = "d9rh5vcywjgurbdxx7kbdqnnybu"
;~ Global Const $sDBID = "bfczwdvwp"

;~ Global Const $dbid = "bfczwdvv6"

; ================================================================================
; ----------------------------------------------------- LOCATION / DEVICE ID
; ================================================================================

Global Const $QB_TOKEN_THAM_HEALTHCHECK = "d9rh5vcywjgurbdxx7kbdqnnybu"
Global Const $QB_DBID_THAM_HEALTHCHECK = "bfczwdvwp"

Global Const $QB_BLDG = 78
Global Const $QB_FLOOR = 79
Global Const $QB_ROOM = 80
Global Const $QB_CONTACTNAME = 13
Global Const $QB_CONTACTPHONE = 14
Global Const $QB_DEPT = 12
Global Const $QB_MOUNTTYPE = 10

Global Const $QB_RID = 3
Global Const $QB_DEVICEMAKE = 6
Global Const $QB_DEVICEMODEL = 7
Global Const $QB_DEVICESERIAL = 9
Global Const $QB_DEVICEASSET = 8
Global Const $QB_DEVICENAME = 16
Global Const $QB_DEVICENEWNAME = 118

Global Const $QB_CLINICAL = 17
Global Const $QB_CLINICALID = 15

; ================================================================================
; ----------------------------------------------------- MAIN SOFTWARE
; ================================================================================

Global Const $QB_OU_CHECKED = 18
Global Const $QB_GROUP_CHECKED = 19
Global Const $QB_BACKGROUND_CHECKED = 20

Global Const $QB_NIC_CHECKED = 21
Global Const $QB_PASSWORD_CHECKED = 23
Global Const $QB_IETEMP_DELETED = 24
Global Const $QB_TEMP_DELETED = 25

Global Const $QB_LANDESK_CHECKED = 46
Global Const $QB_INVENTORYSCAN = 47
Global Const $QB_ACCESSANYWARE_CHECKED = 48
Global Const $QB_INVISION_CHECKED = 93
Global Const $QB_SSO_CHECKED = 49
Global Const $QB_MIVIEWER_CHECKED = 50
Global Const $QB_SYNAPSE_CHECKED = 51
Global Const $QB_EPIC_CHECKED = 52
Global Const $QB_PNAGENT_CHECKED = 84
Global Const $QB_APPV_CHECKED = 75
Global Const $QB_SCCM_CHECKED = 76
Global Const $QB_SCCM_UPDATED = 83
Global Const $QB_MEDITECH_CHECKED = 77
Global Const $QB_MCAFEE_CHECKED = 53
Global Const $QB_IE7_CHECKED = 71

; ================================================================================
; ----------------------------------------------------- APPV CHECKS
; ================================================================================

Global Const $QB_APPVBAD = 107
Global Const $QB_APPVBAD_AA = 102
Global Const $QB_APPVBAD_MIVIEW = 104
Global Const $QB_APPVBAD_STEDMANS = 105
Global Const $QB_APPVBAD_MICROMEDEX = 106

Global Const $QB_APPVBAD_INV = 103

; ================================================================================
; ----------------------------------------------------- MISC HARDWARE / WIRELESS
; ================================================================================

Global Const $QB_PRINTER_CHECKED = 54
Global Const $QB_DEFAULTPRINTER = 55

Global Const $QB_OLDBIOS = 26
Global Const $QB_NEWBIOS = 27

Global Const $QB_HPBIOS_CHECKED = 28

Global Const $QB_WIRELESS_SSID = 29
Global Const $QB_WIRELESS_SIGNAL = 30
Global Const $QB_WIRELESS_NOISE = 31
Global Const $QB_WIRELESS_CHANNEL = 32
Global Const $QB_WIRELESS_BSSID = 33

Global Const $QB_CART_TYPE = 34
Global Const $QB_CART_MODEL = 35
Global Const $QB_CART_ASSET = 36
Global Const $QB_CART_PLUGGED = 37
Global Const $QB_CART_POWER = 38

Global Const $QB_KEYBOARD_COVER = 86
Global Const $QB_MOUNT_ADJUSTED = 87
Global Const $QB_PILLTRAY_INSTALLED = 90
Global Const $QB_SCANNER_INSTALLED = 91
Global Const $QB_SCANNER_TESTED = 92
Global Const $QB_PRIVACYSCREEN = 94

; ================================================================================
; ----------------------------------------------------- MONITOR REPLACEMENT
; ================================================================================

Global Const $QB_OLDMONITOR_SIZE = 95 ; values: CRT | 15" LCD | 17" LCD | 19" LCD | >19" LCD
Global Const $QB_OLDMONITOR_ASSET = 98
Global Const $QB_OLDMONITOR_SERIAL = 99
Global Const $QB_OLDMONITOR_REPLACE = 96

Global Const $QB_NEWMONITOR_INSTALL = 97
Global Const $QB_NEWMONITOR_ASSET = 100
Global Const $QB_NEWMONITOR_SERIAL = 101

; ================================================================================
; ----------------------------------------------------- CLINICAL
; ================================================================================

Global Const $QB_POWER_CHECKED = 39
Global Const $QB_POWER_MAXBATTERY = 40
Global Const $QB_POWER_MONITOR = 41
Global Const $QB_POWER_DISK = 42
Global Const $QB_POWER_HIBERNATION = 43
Global Const $QB_POWER_PASSWORD = 44
Global Const $QB_POWER_ASK = 45

; ================================================================================
; ----------------------------------------------------- ISSUES CHECKBOXES
; ================================================================================

Global Const $QB_ISSUE_REVISIT = 62
Global Const $QB_ISSUE_KBCOVER = 122
Global Const $QB_ISSUE_PSCREEN = 123
Global Const $QB_ISSUE_MONITOR = 119
Global Const $QB_ISSUE_CONVERTCLINICAL = 132
Global Const $QB_ISSUE_CONVERTCOMMON = 133
Global Const $QB_ISSUE_OFFLINE = 135

Global Const $QB_ISSUE_APPV = 124
Global Const $QB_ISSUE_SCCM = 125
Global Const $QB_ISSUE_CC = 130
Global Const $QB_ISSUE_MEDITECH = 126
Global Const $QB_ISSUE_NAME = 127
Global Const $QB_ISSUE_ID = 128
Global Const $QB_ISSUE_STORED = 136

Global Const $QB_ISSUE_JACK = 113
Global Const $QB_ISSUE_POWER = 112
Global Const $QB_ISSUE_HARDWARE = 120
Global Const $QB_ISSUE_REIMAGE = 129
Global Const $QB_ISSUE_OTHER = 121
Global Const $QB_ISSUE_MISSING = 134
Global Const $QB_ISSUE_NOTINSTALLED = 137

Global Const $QB_REVISIT_APPV_RESOLVED = 109
Global Const $QB_REVISIT_APPV_TECH = 110
Global Const $QB_REVISIT_APPV_DATE = 111

Global Const $QB_REVISIT_RESOLVED = 114
Global Const $QB_REVISIT_TECH = 115
Global Const $QB_REVISIT_DATE = 116

; ================================================================================
; ----------------------------------------------------- VERIFICATIONS
; ================================================================================

Global Const $QB_DATE = 57
Global Const $QB_TEAM = 63
Global Const $QB_TECH = 64
Global Const $QB_LOCATED = 72
Global Const $QB_100 = 60
Global Const $QB_ISSUES_CORRECTED = 61
Global Const $QB_COMMENTS = 56

#cs
;~ ====================================================
;~ -- AMH Deployments
;~ ====================================================


GLOBAL CONST $QB_CONTACT = 18
Global Const $QB_DEPT = 17
GLOBAL CONST $QB_DATE = 63

GLOBAL CONST $QB_BUILDING = 43
GLOBAL CONST $QB_FLOOR = 44
GLOBAL CONST $QB_ROOM = 45

Global Const $QB_CLINICAL = 30
GLOBAL CONST $QB_CLINICAL_LOGIN = 215
GLOBAL CONST $QB_DEVICE_NAME = 206
Global Const $QB_DEVICE_MAC = 34
Global Const $QB_DEVICE_WIRELESS = 57

GLOBAL CONST $QB_DEVICE_SERIAL = 24
GLOBAL CONST $QB_DEVICE_ASSET = 25
GLOBAL CONST $QB_DEVICE_MODEL = 79

GLOBAL CONST $QB_MON_SERIAL = 26
GLOBAL CONST $QB_MON_ASSET = 27
Global Const $QB_MON_SETUP = 148 ; Laptop - Monitor integrated / Dual monitor - 2 new monitors

Global Const $QB_MOUNT_TYPE = 80
GLOBAL CONST $QB_MOUNT_SERIAL = 28
GLOBAL CONST $QB_MOUNT_ASSET = 29
GLOBAL CONST $QB_MOUNT_DATE = 64

Global Const $appToken = "b4ick8cdbk9y3ch4c6r5djztjd5"
Global Const $dbid = "be7achjda"
#ce

Global Const $QB_SM_UPDATED = 75
GLOBAL CONST $QB_SM_NOTES = 224

Func _QBVerify($sDBID, $sAppToken, $ticket, $key, $date, $comments = "" )
	$sVerifyURL = "https://dell.quickbase.com/db/" & $sDBID & _
	"?act=API_EditRecord&" & _
	"appToken=" & $sAppToken & _
	"&ticket=" & $ticket & _
	"&key=" & $key & _
	"&_fid_117=" & _QBStringToDate(_NowDate())

	if $comments Then
		$sVerifyURL &= "&_fid_118=" & $comments
	EndIf

	$response = _INetGetSource($sVerifyURL)

	_XMLLoadXML($response)
	$errCode = _XMLGetValue("/qdbapi/errcode")
	$errText = _XMLGetValue("/qdbapi/errtext")

	if $errCode[1] <> 0 Then
		SetError(1)
		SetExtended($errCode[1])
		Return 0
	EndIf

	$iChanged = _XMLGetValue("/qdbapi/num_fields_changed")

	Return $iChanged[1]
EndFunc

Func _QBUpdateField($sDBID, $sAppToken, $ticket, $rID, $fieldID, $contents)
	$sUpdateURL = "https://dell.quickbase.com/db/" & $sDBID & _
	"?act=API_EditRecord&" & _
	"appToken=" & $sAppToken & _
	"&ticket=" & $ticket & _
	"&key=" & $rID & _
	"&_fid_" & $fieldID & "=" & $contents

	$response = _INetGetSource($sUpdateURL)

	_XMLLoadXML($response)
;~ 	if $debug Then ConsoleWrite("QB Debug: " & $response & @CRLF & "----------" & @CRLF)
	$errCode = _XMLGetValue("/qdbapi/errcode")
	$errText = _XMLGetValue("/qdbapi/errtext")

	if $errCode[1] <> 0 Then
		Return SetError(1, $errCode[1])
	EndIf

	$iChanged = _XMLGetValue("/qdbapi/num_fields_changed")

	Return $iChanged[1]
EndFunc

Func _QBGetField($sDBID, $sAppToken, $ticket, $rID, $fieldID = -1)
	$sQueryURL = "https://dell.quickbase.com/db/" & $sDBID & _
	"?act=API_GetRecordInfo&" & _
	"appToken=" & $sAppToken & _
	"&ticket=" & $ticket & _
	"&rid=" & $rID

	$response = _INetGetSource($sQueryURL)
	_XMLLoadXML($response)
	if @error Then Return SetError(2, @error)

	$errCode = _XMLGetValue("/qdbapi/errcode")
	$errText = _XMLGetValue("/qdbapi/errtext")

;~ 	ConsoleWrite($response)

	if $errCode[1] <> 0 Then
		ConsoleWrite("_QBGetField() ERROR: " & $errCode[1] & " Description: " & $errText[1] & @CRLF)
		SetError(1)
		Return 0
	EndIf

	if $fieldID <> -1 Then
		$value = _XMLGetValue("/qdbapi/field[fid=" & $fieldID & "]/value")
		if @error Then
			ConsoleWriteError("getfield error " & @error & ", " & @extended & @CRLF)
			Return SetError(1)
		Else
			Return $value[1]
		EndIf
	Else
		$iFields = _XMLGetValue("/qdbapi/num_fields")

		Dim $aReturn[$iFields[1]+1][3]
		$aReturn[0][0] = $iFields[1]

		$aTemp = _XMLGetValue("/qdbapi/field/fid")
		For $x=1 To $aTemp[0]
			$aReturn[$x][0] = $aTemp[$x]
		Next

		$aTemp = _XMLGetValue("/qdbapi/field/name")
		For $x=1 To $aTemp[0]
			$aReturn[$x][1] = $aTemp[$x]
		Next

		$aTemp = _XMLGetValue("/qdbapi/field/value")
		For $x=1 To $aTemp[0]
			$aReturn[$x][2] = $aTemp[$x]
		Next

		Return $aReturn
	EndIf
EndFunc

Func _QBQuery($sDBID, $sAppToken, $ticket, $sQuery, $cListFields = "", $bStructured = False)
	; build query string
	$sQueryURL = "https://dell.quickbase.com/db/" & $sDBID & _
	"?act=API_DoQuery&" & _
	"appToken=" & $sAppToken & _
	"&ticket=" & $ticket & _
	"&query=" & $sQuery & _
	"&cList=" & $cListFields

	if $bStructured Then _
		$sQueryURL &= "&fmt=structured"

	$response = _INetGetSource($sQueryURL)
	_XMLLoadXML($response)

	; check for error
	$errorCode = _XMLGetValue("/qdbapi/errcode")
	$errorText = _XMLGetValue("/qdbapi/errtext")
	if $errorCode[1] <> 0 Then
		ConsoleWrite("_QBQuery() ERROR: " & $errorCode[1] & " Description: " & $errorText[1] & @CRLF)
		SetError(2)
		SetExtended($errorCode[1])
		Return 0
	EndIf

	;response stuff we care about
	$aFields = _XMLGetValue("/qdbapi/record/record_id_")
	if NOT IsArray($aFields) Then Return SetError(1)

	Return $aFields
EndFunc

Func _QBQueryCount($sDBID, $sAppToken, $ticket, $sQuery)
	; build query string, ie
	; https://dell.quickbase.com/db/benpwsajs?act=API_DoQueryCount&
	; appToken=ba4yarybftvryjbh43pnpcrudfzf&
	; ticket=5_beyre8xhc_bydb7m_uge_dgyfvfh9z4nuwdu2jkvdcyc4f3_b737vy85rv2ge_&
	; query={'SERIALFIELD'.CT.'SERIAL'}OR{'ASSETFIELD'.CT.'SERIAL'}
	$sQueryURL = "https://dell.quickbase.com/db/" & $sDBID & _
	"?act=API_DoQueryCount&" & _
	"appToken=" & $sAppToken & _
	"&ticket=" & $ticket & _
	"&query=" & $sQuery

;~ 	$sQueryURL &= "&clist=3"

	$response = _INetGetSource($sQueryURL)
	_XMLLoadXML($response)

	; check for error
	$errorCode = _XMLGetValue("/qdbapi/errcode")
	$errorText = _XMLGetValue("/qdbapi/errtext")
	if $errorCode[1] <> 0 Then
		ConsoleWrite("_QBQueryCount() ERROR: " & $errorCode[1] & " Description: " & $errorText[1] & @CRLF)
		SetError(1)
		;SetExtended($errorCode[1])
		Return 0
	EndIf

	$iCount = _XMLGetValue("/qdbapi/numMatches")
;~ 	$iCount = _XMLGetValue("/qdbapi/record/rounding_record_id_")
	if IsArray($iCount) Then
		Return $iCount[1]
	Else
		Return 0
	EndIf

EndFunc

Func _QBRecordExists($sDBID, $sAppToken, $ticket, $rID)
	$sQueryURL = "https://dell.quickbase.com/db/" & $sDBID & _
	"?act=API_GetRecordInfo&" & _
	"appToken=" & $sAppToken & _
	"&ticket=" & $ticket & _
	"&rid=" & $rID

	$response = _INetGetSource($sQueryURL)
	_XMLLoadXML($response)

	$errorCode = _XMLGetValue("/qdbapi/errcode")
	$errorText = _XMLGetValue("/qdbapi/errtext")

	if $errorCode[1] <> 0 Then
		SetError(1)
		SetExtended($errorCode[1])
		Return 0
	Else
		Return 1
	EndIf
EndFunc

Func _QBGetApps($ticket)
	$sURL = "https://dell.quickbase.com/db/main?act=API_GrantedDBs&" & _
	"ticket=" & $ticket

	$response = _INetGetSource($sURL)

	_XMLLoadXML($response)

	$dbName = _XMLGetValue("/qdbapi/databases/dbinfo/dbname")
	$sDBID2 = _XMLGetValue("/qdbapi/databases/dbinfo/dbid")

	Dim $arrReturn[$dbName[0]][2]

	For $i=1 To $dbName[0]-1
		$arrReturn[$i][0] = $dbName[$i]
		$arrReturn[$i][1] = $sDBID2[$i]
	Next

	Return $arrReturn
EndFunc

Func _QBAuthGUI()

	$frmAuth = GUICreate("_QBAuth", 185, 174, 259, 391)
	$inAuthUser = GUICtrlCreateInput("", 32, 34, 121, 21)
	$inAuthPass = GUICtrlCreateInput("", 32, 82, 121, 21, BitOR($ES_PASSWORD,$ES_AUTOHSCROLL))
	GUICtrlCreateLabel("Username:", 32, 16, 55, 17)
	GUICtrlCreateLabel("Password:", 32, 64, 53, 17)
	$btnAuthCancel = GUICtrlCreateButton("&Cancel", 96, 136, 75, 25)
	$btnAuthOK = GUICtrlCreateButton("&OK", 16, 136, 75, 25, $BS_DEFPUSHBUTTON)
	GUISetState(@SW_SHOW, $frmAuth)

	While 1
		Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			GUIDelete($frmAuth)
			Return SetError(1)
		Case $btnAuthCancel
			GUIDelete($frmAuth)
			Return SetError(1)
		Case $btnAuthOK
			GUISetState(@SW_DISABLE, $frmAuth)

			$aReturn = _QBAuth(GUICtrlRead($inAuthUser), GUICtrlRead($inAuthPass))

			if @error Then
				MsgBox(48, "Could not authenticate", "Code: " & @error & @CRLF & $aReturn)
				GUISetState(@SW_ENABLE, $frmAuth)
				WinActivate($frmAuth)
			Else
				GUIDelete($frmAuth)
				Return $aReturn
			EndIf
		EndSwitch
		Sleep(10)
	WEnd
EndFunc

Func _QBAuth($user, $pass)
	$targetDomain = "https://dell.quickbase.com/db/main"
	$act = "API_Authenticate"
	$hours = "1"
	$url = $targetDomain & "?act=" & $act & "&username=" & $user & "&password=" & $pass & "&hours=" & $hours
	$response = _INetGetSource($url)
	if StringInStr($response, "Bad Request") Then
		Return SetError(99)
	EndIf

	_XMLLoadXML($response)
	if @error Then
		Return SetError(@error)
	EndIf

	if $debug Then
		ConsoleWrite($response)
	EndIf

	$errCode = _XMLGetValue("/qdbapi/errcode")
	$errText = _XMLGetValue("/qdbapi/errtext")

	if $errCode[1] <> 0 Then
		SetError($errCode[1])
		Return $errText[1]
	EndIf

	$ticket = _XMLGetValue("/qdbapi/ticket")
	$userID = _XMLGetValue("/qdbapi/userid")

	Dim $arrReturn[2] = [$userID[1], $ticket[1]]
	Return $arrReturn
EndFunc

Func _QBSignOut()
	$response = _INetGetSource("https://dell.quickbase.com/db/main?act=API_SignOut")
	_XMLLoadXML($response)
	$return = _XMLGetValue("/qdbapi/errcode")

	Return $return[1]
EndFunc

Func _QBDateToString($iDate)
	$temp = _DateAdd("s", $iDate/1000, "1970/01/01")
	$return = StringRight($temp, 5) & "/" & StringLeft($temp, 4)
	Return $return
EndFunc

Func _QBStringToDate($sDate)
	$sDate = StringRight($sDate, 4) & "/" & StringLeft($sDate, 5)
	$iDate = _DateDiff("s", "1970/01/01", $sDate)*1000

	Return $iDate
EndFunc

Func _QBIEShowRecord($dbID, $rID, $iEdit = 0)
	;https://dell.quickbase.com/db/bfczwdvwp?a=dr&rid=241
	if $iEdit Then
		$sURL = "https://dell.quickbase.com/db/" & $dbID & "?a=er&rid=" & $rID
	Else
		$sURL = "https://dell.quickbase.com/db/" & $dbID & "?a=dr&rid=" & $rID
	EndIf

	$oIE_QB = _IECreate($sURL)
	if @error Then
		Return SetError(@error)
	Else
		Return $oIE_QB
	EndIf

EndFunc

Func _QBIEGetLabel(ByRef $oIE, $sSearch)
	if NOT IsObj($oIE) Then Return SetError(1)

	$oForm = _IEFormGetObjByName($oIE, "editform")
	$oFields = $oForm.All.Tags("td")

	For $oField in $oFields
		if StringInStr($oField.InnerText, $sSearch) AND StringLeft($oField.ID, 4) = "tdl_" Then
			ConsoleWrite($oField.InnerText & @CRLF)
			Return $oField.ID
		EndIf

	Next
EndFunc

Func _QBIEColorFields($sFields, $hColor)
	$oIE = _IEAttach("Sign In")
	if NOT @error Then
		While 1
			Sleep(500)
			if IsObj($oIE) Then
				$sProp = _IEPropertyGet($oIE, "title")
				if @error Then Return SetError(@error)

				if $sProp <> "Sign In" Then ExitLoop
			Else
				Return SetError(1)
			EndIf
		WEnd
		Sleep(500)
		_IELoadWait($oIE)
	Else
		$oIE = _IEAttach("Edit Health Check Record")
		if @error Then
			Return SetError(1)
		EndIf

	EndIf

	$oForm = _IEFormGetObjByName($oIE, "editform")

	$aFields = StringSplit($sFields, ".")
	if $debug Then ConsoleWrite("Total fields: " & $aFields[0] & @CRLF)

	For $i=1 To $aFields[0]
		$oElement = _IEFormElementGetObjByName($oForm, "_fid_" & $aFields[$i])
		if NOT @error Then _
			$oElement.Style.BackgroundColor = $hColor
		if $debug Then ConsoleWrite("Field: _fid_" & $aFields[$i] & @CRLF)
	Next

	Return 1
EndFunc


#cs
Func OnAutoItExit()
	ConsoleWrite("Signing out of QuickBase... ")
	_QBSignOut()
	ConsoleWrite("done." & @CRLF)
EndFunc
#ce