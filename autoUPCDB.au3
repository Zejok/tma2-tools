;~ #AutoIt3Wrapper_aut2exe=C:\Program Files (x86)\AutoIt3\aut2exe\aut2exe.exe
#AutoIt3Wrapper_UseX64=N
#AutoIt3Wrapper_icon=Icons\TMA2\autoUPCDB.ico
#AutoIt3Wrapper_Res_Fileversion=0.1.1.5
#AutoIt3Wrapper_res_fileversion_autoincrement=Y
#AutoIt3Wrapper_res_icon_add=Icons\autoUPCDB.ico
#AutoIt3Wrapper_res_icon_add=Icons\Gnome\16\Dialog-Error.ico
#AutoIt3Wrapper_res_icon_add=Icons\Gnome\emblem-default.ico

#cs -- UPCDB XMLRPC REFERENCE --
help - returns string
	Show available functions and their parameters.

lookupEAN(ean string) - returns struct
	Lookup upc database entry.

lookupUPC(upc string) - returns struct
	Lookup upc database entry.
	DEPRECATED!  Use 'lookupEAN' instead.

writeEntry(username string, password string, ean string, description string, pkgsize string)
	Add or modify an entry in the database.
	This function is unimplemented at this time,
	but is here for API review.

calculateCheckDigit(partialean string) - returns string
	Parameter 'ean' should have 'C' or 'X' in
	place of the check digit (last character).
	Length of 'partialean' parameter should be
	11 or 12 digits, plus 'X' or 'C' character.

convertUPCE(upce string) - returns string
	Parameter 'upce' should be exactly 8 digits.
	Returns full EAN-13.

decodeCueCat(cuecatscan string) - returns struct
	Returns serial number, type, and code given
	CueCat scanner output.

latestDownloadURL - returns string
	Return URL of latest full database download.

------
Return values:

upc					string ex: 012345678912
pendingUpdates		bool
isCoupon			bool
ean					string ex: 0123456789123
issuerCountryCode	string ex: us
found				bool
description			string ex: box of 12 buttplugs, industrial grade
size				string ex: 13lb
message				string ex: Database entry found
issuerCountry		string ex: United States
LastModified		string ex: YYYY-MM-DD HH:MM:SS

#ce

if NOT RegRead("HKCR\vbXMLRPC.XMLRPCRequest", "") Then
	$q = MsgBox(4100, "Error", 'This requires vbXMLRPC. Please download it and run "regsvr32 <path to vbXMLRPC.dll>"' & @CRLF & @CRLF & "Visit site now?")
	if $q = 6 Then
		$sBrowser = RegRead("HKCU\Software\Classes\http\shell\open\command", "")

		$sBrowserPath = StringTrimLeft(StringLeft($sBrowser, StringInStr($sBrowser, '"', 0, 2)), 1)
		Debug_autoUPC($sBrowserPath)

		ShellExecute($sBrowserPath, "http://www.enappsys.com/backend/vbXMLRPC/vbXMLRPCBinaries.jsp")
	EndIf
	Exit 1
EndIf

Global $oError = ObjEvent("AutoIt.Error", "user_error")
Global $verbose = 0

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <guiListView.au3>
#include <guiImageList.au3>

#include <modernMenuRaw.au3>

#Region ### START Koda GUI section ### Form=C:\Users\TMA2\Documents\Scripts\autoUPCDB.kxf
$Form1 = GUICreate("autoUPCDB", 380, 338, 408, 407, BitOR($WS_MINIMIZEBOX,$WS_SIZEBOX))

$label1 = GUICtrlCreateLabel("Scan EAN:", 8, 8, 57, 17)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
$Input1 = GUICtrlCreateInput("", 72, 8, 230, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_NUMBER))
GUICtrlSetResizing(-1, $GUI_DOCKHEIGHT+$GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKRIGHT)
GUICtrlSetLimit(-1, 13)
$Button1 = GUICtrlCreateButton("Go", 309, 8, 59, 21, $BS_DEFPUSHBUTTON)
GUICtrlSetResizing(-1, $GUI_DOCKSIZE+$GUI_DOCKTOP+$GUI_DOCKRIGHT)
$ListView1 = GUICtrlCreateListView("EAN|Item|Size|Status", 0, 40, 378, 252, BitOR($LVS_REPORT,$LVS_EDITLABELS), BitOR($LVS_EX_DOUBLEBUFFER,$LVS_EX_HEADERDRAGDROP,$LVS_EX_FULLROWSELECT))
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)

GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 100)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 150)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 2, 50)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 3, 80)

$menu = GUICtrlCreateContextMenu($ListView1)
;$menuAdd = GUICtrlCreateMenuItem("&Add", $menu)
$menuAdd = _GUICtrlCreateODMenuItem("Add" & @TAB & "Ctrl-A", $menu, "Icons\Gnome\16\List-Add.ico")
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlCreateMenuItem("", $menu)
;$menuCopy = GUICtrlCreateMenuItem("&Copy", $menu)
$menuCopy = _GUICtrlCreateODMenuItem("Copy" & @TAB & "Ctrl-C", $menu, "Icons\Gnome\16\Edit-Copy.ico")
$menuPaste = _GUICtrlCreateODMenuItem("Paste" & @TAB & "Ctrl-V", $menu, "Icons\Gnome\16\Edit-Paste.ico")
;$menuCSV = GUICtrlCreateMenuItem("&Save (CSV)", $menu)
$menuCSV = _GUICtrlCreateODMenuItem("Save to CSV" & @TAB & "Ctrl-S", $menu, "Icons\Gnome\16\stock_data-save.ico")
GUICtrlCreateMenuItem("", $menu)
;$menuClear = GUICtrlCreateMenuItem("C&lear", $menu)
$menuClear = _GUICtrlCreateODMenuItem("Clear Selected" & @TAB & "Del", $menu, "Icons\Gnome\16\stock_delete-row.ico")
$menuClearA = _GUICtrlCreateODMenuItem("Clear All" & @TAB & "Ctrl-Del", $menu, "Icons\Gnome\16\stock_data-delete-table.ico")

;~ GUICtrlCreateListViewItem("temp", $ListView1)
;~ GUICtrlCreateListViewItem("temp", $ListView1)
;~ GUICtrlCreateListViewItem("temp", $ListView1)
;~ _GUICtrlListView_AddItem($ListView1, "temp", 0, 58)
;~ _GUICtrlListView_AddItem($ListView1, "temp", 0, 59)
;~ _GUICtrlListView_AddItem($ListView1, "temp", 0, 60)

_GUICtrlListView_EnableGroupView($ListView1)
_GUICtrlListView_InsertGroup($ListView1, -1, 0, "Present")
_GUICtrlListView_InsertGroup($ListView1, -1, 1, "Missing")
_GUICtrlListView_InsertGroup($ListView1, -1, 2, "Invalid")

_GUICtrlListView_SetGroupInfo($ListView1, 0, "Present", $LVGS_COLLAPSIBLE)
_GUICtrlListView_SetGroupInfo($ListView1, 1, "Missing", BitOR($LVGS_COLLAPSIBLE,$LVGS_SUBSETED))
_GUICtrlListView_SetGroupInfo($ListView1, 2, "Invalid", BitOR($LVGS_COLLAPSIBLE,$LVGS_COLLAPSED,$LVGS_SUBSETED))

_GUICtrlListView_SetItemGroupID($ListView1, 0, 0)
_GUICtrlListView_SetItemGroupID($ListView1, 1, 1)
_GUICtrlListView_SetItemGroupID($ListView1, 2, 2)

$StatusBar1 = _GUICtrlStatusBar_Create($Form1)
_GUICtrlStatusBar_SetSimple($StatusBar1)

$hImages = _GUIImageList_Create(16, 16, 5)
_GUICtrlListView_SetImageList($ListView1, $hImages, 1)

if @Compiled Then
	_GUIImageList_AddIcon($hImages, @ScriptFullPath, 6)
	_GUIImageList_AddIcon($hImages, @ScriptFullPath, 5)
Else
	_GUIImageList_AddIcon($hImages, "Icons\Gnome\16\emblem-default.ico")
	_GUIImageList_AddIcon($hImages, "Icons\Gnome\16\Dialog-Error.ico")
EndIf


Local $aAccel[6][2] = [["^a", $menuAdd], ["^c", $menuCopy], ["^s", $menuCSV], ["{delete}", $menuClear], ["^{delete}", $menuClearA], ["^v", $menuPaste]]
GUISetAccelerators($aAccel)

GUIRegisterMsg($WM_SIZE, "WM_SIZE")

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Button1
			$sIn = GUICtrlRead($Input1)

			if StringLen($sIn) = 12 Then
				$sIn = "0" & $sIn
				GUICtrlSetData($Input1, $sIn)
			ElseIf StringLen($sIn) < 12 Then
				_GUICtrlStatusBar_SetText($StatusBar1, "Please enter a 12 (UPC) or 13 (EAN) character string.")
				ContinueLoop
			EndIf

			_gui_checkEAN($sIn)
		Case $menuClearA
			_GUICtrlListView_DeleteAllItems($ListView1)
		Case $menuClear
			_GUICtrlListView_DeleteItemsSelected($ListView1)
		Case $menuPaste
			$hFocus = ControlGetHandle($Form1, "", ControlGetFocus($Form1))
			if $hFocus = GUICtrlGetHandle($Input1) Then
				GUICtrlSetData($Input1, ClipGet())
				ContinueLoop
			ElseIf $hFocus <> GUICtrlGetHandle($ListView1) Then
				ContinueLoop
			EndIf

			$aValues = _stringParseUPCs(ClipGet())
			if @error Then
				_GUICtrlStatusBar_SetText($StatusBar1, "No UPCs found in clipboard.")
				ContinueLoop
			EndIf

			$ghTimer = TimerInit()
			For $x=1 To $aValues[0]
				if StringLen($aValues[$x]) = 12 Then $aValues[$x] = "0" & $aValues[$x]
				_gui_checkEAN($aValues[$x])
			Next

			_GUICtrlStatusBar_SetText($StatusBar1, $aValues[0] & " items queried in " & Round(TimerDiff($ghTimer)/1000, 2) & "s.")
		Case $menuCopy
			Local $aSelected = _GUICtrlListView_GetSelectedIndices($ListView1, True)
			Local $sCopy = ""

			For $index=1 To $aSelected[0]
				$sCopy &= _GUICtrlListView_GetItemText($ListView1, $aSelected[$index]) & " - "
				$sCopy &= _GUICtrlListView_GetItemText($ListView1, $aSelected[$index], 1) & " - "
				if _GUICtrlListView_GetItemText($ListView1, $aSelected[$index], 2) Then _
					$sCopy &= _GUICtrlListView_GetItemText($ListView1, $aSelected[$index], 2) & " - "
				$sCopy &= @CRLF
			Next

			ClipPut($sCopy)
		Case $menuCSV
			if _GUICtrlListView_GetItemCount($ListView1) = 0 Then ContinueLoop

			$sPath = FileSaveDialog("Save (CSV)", @MyDocumentsDir, "All (*.*)", 0, "autoUPCDB-" & @UserName & ".csv", $Form1)
			if @error Then ContinueLoop

			Local $sCopy = ""
			if NOT FileExists($sPath) Then FileWriteLine($sPath, "EAN,Description,Size")
			$hFile = FileOpen($sPath, 1)

			For $index=0 To _GUICtrlListView_GetItemCount($ListView1)-1
				$sCopy = ""
				$sCopy &= _GUICtrlListView_GetItemText($ListView1, $index) & ","
				$sCopy &= _GUICtrlListView_GetItemText($ListView1, $index, 1) & ","
				$sCopy &= _GUICtrlListView_GetItemText($ListView1, $index, 2)
				$sCopy &= @CRLF
				FileWrite($hFile, $sCopy)
			Next

			FileClose($hFile)
		Case $menuAdd
			$aSelection = _GUICtrlListView_HitTest(GUICtrlGetHandle($ListView1))
			if IsArray($aSelection) Then
				if _GUICtrlListView_GetItemText($ListView1, $aSelection[0], 3) <> "OK" Then
					ShellExecute("http://www.upcdatabase.com/addform.asp?upc=" & _GUICtrlListView_GetItemText($ListView1, $aSelection[0]))
				EndIf
			Else
				Debug_autoUPC("No item selected")
			EndIf
		Case $ListView1
			#cs
			$aSelection = _GUICtrlListView_HitTest(GUICtrlGetHandle($ListView1))
			if IsArray($aSelection) Then
				if _GUICtrlListView_GetItemText($ListView1, $aSelection[0], 3) = "OK" Then
					GUICtrlSetState($menuAdd, $GUI_DISABLE)
				Else
					GUICtrlSetState($menuAdd, $GUI_ENABLE)
				EndIf
			EndIf
			#ce
	EndSwitch
WEnd

Func _UPCDB_LookupEAN($sEAN)
	Local Const $XMLRPC_NOTINITIALISED = 0
	Local Const $XMLRPC_HTTPERROR = 1
	Local Const $XMLRPC_FAULTRETURNED = 2
	Local Const $XMLRPC_PARAMSRETURNED = 3
	Local Const $XMLRPC_XMLPARSEERROR = 4
	Local $aReturn[3]
	$oXMLRPCRequest = ObjCreate("vbXMLRPC.XMLRPCRequest")

	if @error Then
		$iTemp = @error
		Debug_autoUPC("Error " & @error & " creating XMLRPC object.")
		Return SetError(1, $iTemp)
	EndIf

	$oXMLRPCRequest.HostName = "www.upcdatabase.com"
	$oXMLRPCRequest.HostPort = 80
	$oXMLRPCRequest.HostURI = "/rpc"
	$oXMLRPCRequest.MethodName = "lookupEAN"

	$oXMLRPCRequest.Params.AddString($sEAN)

	$oXMLRPCResponse = $oXMLRPCRequest.Submit()

	if $verbose Then
		Debug_autoUPC("Return: " & @CRLF & $oXMLRPCResponse.Status & @CRLF & _
			$oXMLRPCResponse.XMLResponse & _
			"----------" & @CRLF)
	EndIf

	if $oXMLRPCResponse.Status <> $XMLRPC_PARAMSRETURNED Then
		Switch $oXMLRPCResponse.Status
		Case $XMLRPC_FAULTRETURNED
			SetError(2, $XMLRPC_FAULTRETURNED)
			Return $oXMLRPCResponse.Fault
		Case $XMLRPC_HTTPERROR
			SetError(2, $XMLRPC_HTTPERROR)
			Return $oXMLRPCResponse.HTTPStatusCode
		Case $XMLRPC_XMLPARSEERROR
			SetError(2, $XMLRPC_XMLPARSEERROR)
			Return $oXMLRPCResponse.XMLParseError
		EndSwitch
	EndIf

	if $oXMLRPCResponse.Params(1).ValueType = 3 Then
		SetError(4)
		Return $oXMLRPCResponse.Params(1).StringValue
	EndIf

	For $oMember in $oXMLRPCResponse.Params(1).StructValue
		if $oMember.Name = "found" AND $oMember.Value.BooleanValue = False Then
			SetError(3)
			Return "EAN not found."
		EndIf

		if $oMember.Name = "ean" Then
			$aReturn[0] = $oMember.Value.StringValue
		Elseif $oMember.Name = "description" Then
			$aReturn[1] = $oMember.Value.StringValue
		ElseIf $oMember.Name = "size" Then
			$aReturn[2] = $oMember.Value.StringValue
		EndIf
	Next

	$oXMLRPCRequest = 0
	$oXMLRPCResponse = 0

	Return $aReturn
EndFunc

Func _gui_checkEAN($sEAN)
	$hTimer = TimerInit()

	$iIndex = _GUICtrlListView_FindText($ListView1, $sEAN)
	if $iIndex = -1 Then
		GUICtrlCreateListViewItem($sEAN, $ListView1)
		$iIndex = _GUICtrlListView_GetItemCount($ListView1)-1
;~ 		$iIndex = _GUICtrlListView_AddItem($ListView1, GUICtrlRead($Input1))
	Else
		_GUICtrlListView_SetItemText($ListView1, $iIndex, $sEAN)
	EndIf

	_GUICtrlListView_SetSelectionMark($ListView1, $iIndex)
	_GUICtrlListView_SetColumnWidth($ListView1, 0, $LVSCW_AUTOSIZE)

	$aResponse = _UPCDB_LookupEAN($sEAN)

	if @error Then
		$iErr = @error
		$iExt = @extended

		if $iErr = 3 Then
			_GUICtrlListView_SetItemText($ListView1, $iIndex, "Not found", 3)
			_GUICtrlStatusBar_SetText($StatusBar1, "Item not found. " & Round(TimerDiff($hTimer)/1000, 2) & "s.")
			_GUICtrlListView_SetItemGroupID($ListView1, $iIndex, 1)
			_GUICtrlListView_SetGroupInfo($ListView1, 1, "Missing", BitOR($LVGS_COLLAPSIBLE,$LVGS_SUBSETED))
		ElseIf $iErr = 4 Then
			_GUICtrlStatusBar_SetText($StatusBar1, $aResponse & "... " & Round(TimerDiff($hTimer)/1000, 2) & "s.")
			_GUICtrlListView_SetItemGroupID($ListView1, $iIndex, 2)
			_GUICtrlListView_SetGroupInfo($ListView1, 2, "Invalid", BitOR($LVGS_COLLAPSIBLE,$LVGS_COLLAPSED,$LVGS_SUBSETED))
		ElseIf $iErr = 2 Then
			_GUICtrlListView_SetItemText($ListView1, $iIndex, "Err 2/" & $iExt, 3)
			_GUICtrlListView_SetItemGroupID($ListView1, $iIndex, 2)
		Else
			_GUICtrlStatusBar_SetText($StatusBar1, "Error " & $iErr & "/" & $iExt & ": " & $aResponse)
			_GUICtrlListView_SetItemGroupID($ListView1, $iIndex, 2)
			_GUICtrlListView_SetGroupInfo($ListView1, 2, "Invalid", BitOR($LVGS_COLLAPSIBLE,$LVGS_COLLAPSED,$LVGS_SUBSETED))
		EndIf

		if NOT _GUICtrlListView_SetItemImage($ListView1, $iIndex, 1) Then
			Debug_autoUPC("SetItemImage failed.")
		EndIf

		Return
	EndIf

	_GUICtrlListView_SetItemText($ListView1, $iIndex, $aResponse[1], 1)
	_GUICtrlListView_SetItemText($ListView1, $iIndex, $aResponse[2], 2)
	_GUICtrlListView_SetItemText($ListView1, $iIndex, "OK", 3)

	_GUICtrlListView_SetItemGroupID($ListView1, $iIndex, 0)
	_GUICtrlListView_SetGroupInfo($ListView1, 0, "Present", $LVGS_COLLAPSIBLE)

	_GUICtrlListView_SetColumnWidth($ListView1, 1, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($ListView1, 3, $LVSCW_AUTOSIZE)

	if NOT _GUICtrlListView_SetItemImage($ListView1, $iIndex, 0) Then
		Debug_autoUPC("SetItemImage failed.")
	EndIf

	_GUICtrlStatusBar_SetText($StatusBar1, "Found in " & Round(TimerDiff($hTimer)/1000, 2) & "s.")
EndFunc

Func _stringParseUPCs($s)
	$aResults = StringRegExp($s, "(?:\D)(\d{12,13})(?:\D)", 3)
	$err = @error
	if NOT $err Then
		if $verbose Then Debug_autoUPC("regexp results: " & UBound($aResults))
		ReDim $aResults[UBound($aResults)+1]
		For $x=UBound($aResults)-1 To 1 Step -1
			$aResults[$x] = $aResults[$x-1]
		Next
		$aResults[0] = UBound($aResults)-1

		Return $aResults
	Else
		Debug_autoUPC("_stringParseUPCs error: " & $err)
		Return SetError($err)
	EndIf
EndFunc

Func WM_SIZE($hWnd, $iMsg, $iwParam, $ilParam)
    _GUICtrlStatusBar_Resize ($StatusBar1)
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_SIZE

Func Debug_autoUPC($sMsg, $sLine = @ScriptLineNumber)
	if @Compiled Then
		$sDate = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
		FileWriteLine(@ScriptDir & "\" & @ScriptName & "-" & @ComputerName & "-Debug.log", $sDate & " ... " & $sMsg)
	Else
		ConsoleWrite(@ScriptName & " DEBUG (" & $sLine & ") ----- << " & @CRLF & $sMsg & @CRLF & "     >>" & @CRLF)
	EndIf
EndFunc

Func user_error($sLine = @ScriptLineNumber)
	Debug_autoUPC("ERROR " & $oError.Number & " from " & $oError.Source & ": " & $oError.Description, $sLine)
EndFunc