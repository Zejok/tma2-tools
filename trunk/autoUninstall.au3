; TODO
; re-write function to grab icons for msi apps
; make remote uninstall possible
#AutoIt3Wrapper_Icon=..\Icons\Jon\autoUninstall.ico
#AutoIt3Wrapper_Res_Fileversion=0.3.5.9
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Change2CUI=N

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <WindowsConstants.au3>
#include <array.au3>

#include <guilistview.au3>
#include <guiimagelist.au3>
#include <guimenu.au3>

#include <includes\tma2.au3>
#include <includes\_services.au3>

Opt("TrayIconDebug", 1)

Global $bStop = False
Global $oArgs = ObjCreate("Scripting.Dictionary")
; D11A42B3B10DA2B42BA956A2486831BB
; 3B24A11DD01B4B2AB29A652A848613BB
#region ### START Koda GUI section ### Form=
$Form1 = GUICreate("autoUninstallIt", 339, 422, 566, 408, BitOR($WS_SIZEBOX, $WS_MAXIMIZEBOX, $WS_MINIMIZEBOX))
WinMove($Form1, "", 100, 100, @DesktopWidth / 2, @DesktopHeight / 2)
$ListView1 = GUICtrlCreateListView("Name|Publisher|Version|Size|Install Date|Uninstall String|Location", 0, 0, 338, 358, $LVS_REPORT, BitOR($LVS_EX_DOUBLEBUFFER,$LVS_EX_FULLROWSELECT, $LVS_EX_HEADERDRAGDROP,$LVS_EX_CHECKBOXES))
GUICtrlSetBkColor($ListView1, $GUI_BKCOLOR_LV_ALTERNATE)
GUICtrlSetBkColor($ListView1, 0xDDDDDD)
$lvtemp = GUICtrlCreateListViewItem("temp", $ListView1)
GUICtrlSetBkColor($lvtemp, 0xFFFFFF)
GUICtrlDelete($lvtemp)

if NOT @Compiled Then
	GUISetIcon("Icons\TMA2\autoUninstall.ico")
EndIf

if $cmdLine[0] > 0 Then
	For $i = 1 To $cmdLine[0]
		if StringRegExp($cmdLine[$i], "/\w*:") Then
			$itemp = StringInStr($cmdLine[$i], ":")
			$oArgs.Add(StringMid($cmdLine[$i], 2, $itemp), StringTrimLeft($cmdLine[$i], $itemp))
		EndIf
	Next

	$aitems = $oArgs.Items()
	_ArrayDisplay($aitems)
	$akeys = $oArgs.Keys()
	_ArrayDisplay($akeys)
EndIf


$iCols = _GUICtrlListView_GetColumnCount($ListView1)

Global $hMenuItem[$iCols]

For $i = 0 To $iCols - 1
	$aCol = _GUICtrlListView_GetColumn($ListView1, $i)
	Switch $aCol[5]
		Case "Name"
			Global $iColName = $i
		Case "Publisher"
			Global $iColPub = $i
		Case "Version"
			Global $iColVersion = $i
		Case "Size"
			Global $iColSize = $i
		Case "Install Date"
			Global $iColDate = $i
		Case "Uninstall String"
			Global $iColUninstall = $i
		Case "Location"
			Global $iColLocation = $i
		Case "MSI GUID"
			Global $iColGUID = $i
		Case "Icon Path"
			Global $iColIconPath = $i
	EndSwitch

	$hMenuItem[$i] = 1000 + $i
Next

GUICtrlSetResizing($ListView1, $GUI_DOCKBORDERS)
_GUICtrlListView_JustifyColumn($ListView1, 3, 1)
$lStatus = GUICtrlCreateLabel("Select programs to uninstall", 4, 376, 150, 25)
$Button1 = GUICtrlCreateButton("Uninstall", 256, 368, 75, 25, 0)
;~ _GUIImageList_Create


$hProgress = GUICtrlCreateProgress(180, 370, 60, 20)
GUICtrlSetResizing($hProgress, $GUI_DOCKBOTTOM + $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
GUICtrlSetResizing($lStatus, $GUI_DOCKBOTTOM + $GUI_DOCKLEFT + $GUI_DOCKSIZE)
GUICtrlSetResizing($Button1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

$hImage = _GUIImageList_Create()
_GUICtrlListView_SetImageList($ListView1, $hImage, 1)
_GUIImageList_AddIcon($hImage, @SystemDir & "\shell32.dll", 2)

GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

populateUninstalls()

;~ getSizes()

GUICtrlSetData($lStatus, "Found " & _GUICtrlListView_GetItemCount($ListView1) & " uninstall commands.")
_GUICtrlListView_SetColumnWidth($ListView1, 0, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($ListView1, 1, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($ListView1, 2, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($ListView1, 3, $LVSCW_AUTOSIZE)
_GUICtrlListView_RegisterSortCallBack($ListView1)

;~ $hContext = _GUICtrlMenu_CreatePopup()

GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
;~ GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
;~ GUIRegisterMsg($WM_CONTEXTMENU, "WM_CONTEXTMENU")

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			_GUICtrlListView_UnRegisterSortCallBack($ListView1)
			GUIDelete()
			Exit
		Case $ListView1
			_GUICtrlListView_SortItems($ListView1, GUICtrlGetState($ListView1))
		Case $Button1
			runUninstalls()
	EndSwitch
WEnd

Func runUninstalls($sComp = ".")
	$sItems = ""
	$iCount = _GUICtrlListView_GetItemCount($ListView1)

	GUICtrlSetData($Button1, "Cancel")
	$hTimer = TimerInit()

;~ 	$oWMI = ObjGet("winmgmts://" & $sComp & "/root/cimv2")

	For $i = 0 To $iCount-1
		if _GUICtrlListView_GetItemChecked($ListView1, $i) = False Then ContinueLoop

		_GUICtrlListView_SetSelectionMark($ListView1, $i)
		GUICtrlSetData($lStatus, "Uninstalling " & $i+1 & "/" & $iCount & ": " & _GUICtrlListView_GetItemText($ListView1, $i, 0))
		$iReturn = RunWait(_GUICtrlListView_GetItemText($ListView1, $i, 5))
		ConsoleWrite($i & " return value: " & $iReturn & @CRLF)

		If $iReturn = 0 And Not @error Then
			_GUICtrlListView_DeleteItem($ListView1, $i)
			$i -= 2
		Else
			_GUICtrlListView_SetItemChecked($ListView1, $i, False)
		EndIf

		Sleep(100)
		If $bStop Then
			ExitLoop
		EndIf
	Next
	GUICtrlSetData($Button1, "Uninstall")
	GUICtrlSetData($lStatus, "Done after " & Round(TimerDiff($hTimer) / 1000) & " seconds.")
EndFunc   ;==>runUninstalls

Func populateUninstalls($sComp = @ComputerName)
	Local $iSizeTotal = 0
	Local $aSubkeys, $aSubkeys2
	Local $totalNormalEntries, $totalExtEntries
	Global Const $S_REGLOC64 = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
	Global Const $S_REGLOC = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

	$aSubkeys = _regEnumKeys($S_REGLOC, $sComp)
	if @error Then Return

	$totalNormalEntries = UBound($aSubkeys)
	if StringInStr(@OSArch, "64") Then
		$aSubkeys2 = _regEnumKeys($S_REGLOC64, $sComp)
		$totalExtEntries = UBound($aSubkeys2)
		_ArrayConcatenate($aSubkeys, $aSubkeys2)
	EndIf

	For $i=0 To UBound($aSubkeys)-1
		$sEnum = $aSubkeys[$i]
		if $i < $totalNormalEntries Then
			$sRegLoc = $S_REGLOC
		Else
			$sRegLoc = $S_REGLOC64
		EndIf

		$sKey = "\\" & $sComp & "\HKLM\" & $sRegLoc & "\" & $sEnum

		$sName = RegRead($sKey, "DisplayName")
		If @error Then
			ContinueLoop
		ElseIf RegRead($sKey, "ParentKeyName") Then
			ContinueLoop
		ElseIf RegRead($sKey, "SystemComponent") Then
			ContinueLoop
		EndIf

		$sUninstall = RegRead($sKey, "UninstallString")
		If @error Or $sUninstall = "" Then ContinueLoop

		$sUninstall = StringReplace($sUninstall, "msiexec.exe /i", "msiexec.exe /X")
		If StringInStr($sUninstall, "msiexec.exe") Then $sUninstall &= " /passive"

		$sPub = RegRead($sKey, "Publisher")
		$sVer = RegRead($sKey, "DisplayVersion")
		$sLoc = RegRead($sKey, "InstallLocation")
		$sDate = RegRead($sKey, "InstallDate")
		$sIcon = RegRead($sKey, "DisplayIcon")
		If RegRead($sKey, "EstimatedSize") Then
			$iSize = NumberPadZeroesFloat(Round(Number(RegRead($sKey, "EstimatedSize")) / 1024, 2))
			$iSizeTotal += $iSize
		Else
			$iSize = ""
		EndIf

		GUICtrlSetData($lStatus, $i+1 & " / " & UBound($aSubkeys) & ". " & Round((($i+1)/UBound($aSubkeys))*100) & "% done")
		GUICtrlSetData($hProgress, Round((($i+1)/UBound($aSubkeys))*100))

;~ 		if @error Then ConsoleWrite("regread error: " & @error & @CRLF)
		$hItem = GUICtrlCreateListViewItem($sName, $ListView1)

		$iCurrent = _GUICtrlListView_GetItemCount($ListView1) - 1
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sPub, $iColPub)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sVer, $iColVersion)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $iSize, $iColSize)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sDate, $iColDate)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sUninstall, $iColUninstall)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sLoc, $iColLocation)

		If $sIcon <> "" Then
			$aIcon = StringSplit($sIcon, ",")
			If $aIcon[0] > 1 Then
				$hIcon = _GUIImageList_AddIcon($hImage, $aIcon[1], $aIcon[2])
			Else
				$hIcon = _GUIImageList_AddIcon($hImage, $aIcon[1], 0)
			EndIf
;~ 			_GUICtrlListView_SetItemText($ListView1, $iCurrent, $aIcon[1], $iColIconPath)
			_GUICtrlListView_SetItemImage($ListView1, $iCurrent, $hIcon)
		ElseIf _stringIsGUID($sEnum) Then
			$sNewKey = "HKLM\Software\Classes\Installer\Products\" & _stringConvertUUID($sEnum, 2)

;~ 			_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sNewKey, $iColGUID)
			$sIcon = RegRead("\\" & $sComp & "\" & $sNewKey, "ProductIcon")
			if NOT @error Then
				$hIcon = _GUIImageList_AddIcon($hImage, $sIcon)
				_GUICtrlListView_SetItemImage($ListView1, $iCurrent, $hIcon)
;~ 				_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sIcon, $iColIconPath)
			EndIf
		EndIf
	Next

	_GUICtrlListView_SetColumn($ListView1, $iColSize, "Size : " & Round($iSizeTotal), -1, 1)
	_GUICtrlListView_SetColumnWidth($ListView1, $iColName, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($ListView1, $iColPub, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($ListView1, $iColVersion, $LVSCW_AUTOSIZE)
EndFunc   ;==>populateUninstalls{36A415C2-7181-421D-92C9-8255766E0FF3}

Func getSizes()
	For $i = 0 To _GUICtrlListView_GetItemCount($ListView1) - 1
		Local $sName = _GUICtrlListView_GetItemText($ListView1, $i, $iColName)
		Local $iSize = _GUICtrlListView_GetItemText($ListView1, $i, $iColSize)
		Local $sLoc = _GUICtrlListView_GetItemText($ListView1, $i, $iColLocation)
		If Not $iSize Then
			GUICtrlSetData($lStatus, 'Calculating size of "' & $sName)
			Local $iTotal = 0
			$aSize = _dirSize($iColLocation, $iTotal, $hProgress)
			$iSize = Round($aSize[0] / 1048576)
		EndIf
	Next
EndFunc   ;==>getSizes

Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg, $iwParam, $ilParam
	$tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
	$hWndFrom = DllStructGetData($tNMHDR, "hWndFrom")
	$IDFrom = DllStructGetData($tNMHDR, "IDFrom")
	$Code = DllStructGetData($tNMHDR, "Code")

;~ 	Switch $hWndFrom

EndFunc   ;==>WM_NOTIFY

Func WM_COMMAND($hWnd, $iMsg, $iwParam, $ilParam)
	ConsoleWrite("hWnd:    " & $hWnd & @CRLF)
	ConsoleWrite("iMsg:    " & $iMsg & @CRLF)
	ConsoleWrite("iwParam: " & $iwParam & @CRLF)
	ConsoleWrite("ilParam: " & $ilParam & @CRLF)

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

Func WM_CONTEXTMENU($hWnd, $iMsg, $iwParam, $ilParam)
	Local $hMenu

	$hMenu = _GUICtrlMenu_CreatePopup()

	For $i = 0 To _GUICtrlListView_GetColumnCount($ListView1) - 1
		$aTemp = _GUICtrlListView_GetColumn($ListView1, $i)
		_GUICtrlMenu_AddMenuItem($hMenu, $aTemp[5], $hMenuItem[$i])
	Next

	_GUICtrlMenu_TrackPopupMenu($hMenu, $iwParam)
	_GUICtrlMenu_DestroyMenu($hMenu)
EndFunc   ;==>WM_CONTEXTMENU
