#AutoIt3Wrapper_Icon=..\Icons\Jon\autoUninstall.ico
#AutoIt3Wrapper_Res_Fileversion=0.3.5.12
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

Global $debug = False
Global $bStop = False
Global $oArgs = ObjCreate("Scripting.Dictionary")

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
			$oArgs.Add(StringLeft($cmdLine[$i], $itemp-1), StringTrimLeft($cmdLine[$i], $itemp))
		EndIf
	Next

	if $debug Then
		$aitems = $oArgs.Items()
		_ArrayDisplay($aitems)
		$akeys = $oArgs.Keys()
		_ArrayDisplay($akeys)
	EndIf
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

$hProgress = GUICtrlCreateProgress(180, 370, 60, 20)
GUICtrlSetResizing($hProgress, $GUI_DOCKBOTTOM + $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)
GUICtrlSetResizing($lStatus, $GUI_DOCKBOTTOM + $GUI_DOCKLEFT + $GUI_DOCKSIZE)
GUICtrlSetResizing($Button1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

$hImage = _GUIImageList_Create()
_GUICtrlListView_SetImageList($ListView1, $hImage, 1)
_GUIImageList_AddIcon($hImage, @SystemDir & "\shell32.dll", 2)

GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

if $oArgs.Exists("/rc") Then
	; list software on specified computer
	populateUninstalls($oArgs.Item("/rc"))
Else
	populateUninstalls()
EndIf

;~ getSizes()

GUICtrlSetData($lStatus, "Found " & _GUICtrlListView_GetItemCount($ListView1) & " uninstall commands.")
_GUICtrlListView_SetColumnWidth($ListView1, 0, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($ListView1, 1, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($ListView1, 2, $LVSCW_AUTOSIZE)
_GUICtrlListView_SetColumnWidth($ListView1, 3, $LVSCW_AUTOSIZE)
_GUICtrlListView_RegisterSortCallBack($ListView1)

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

	; cheap, temporary check for 64bit Windows. locally the @OSArch macro can be used, but remotely i'm not sure yet
	Local $osIs64 = False
	RegRead("\\" & $sComp & "\HKLM\Software\Wow6432Node", "")
	if NOT @error Then
		$osIs64 = True
	EndIf

	; populate array with standard \Uninstall keys (at the same time getting total count)
	$aSubkeys = _regEnumKeys($S_REGLOC, $sComp)
	if @error Then Return

	$totalNormalEntries = UBound($aSubkeys)
	if $osIs64 Then
		; add secondary 64-bit software entries if needed
		$aSubkeys2 = _regEnumKeys($S_REGLOC64, $sComp)
		$totalExtEntries = UBound($aSubkeys2)
		_ArrayConcatenate($aSubkeys, $aSubkeys2)
	EndIf

	For $i=0 To UBound($aSubkeys)-1
		$sEnum = $aSubkeys[$i]
		; cheap check if 64-bit registry path needs to be switched to.
		; basically, if the current index is past the total count of 32-bit apps
		if $i < $totalNormalEntries Then
			$sRegLoc = $S_REGLOC
		Else
			$sRegLoc = $S_REGLOC64
		EndIf

		; construct key string
		$sKey = "\\" & $sComp & "\HKLM\" & $sRegLoc & "\" & $sEnum

		$sName = RegRead($sKey, "DisplayName")
		If @error Then ; no displayname? toss that shit
			ContinueLoop
		ElseIf RegRead($sKey, "ParentKeyName") Then
			; if there's a ParentKeyName it's almost always some hotfix or crap inflating
			; the app list, e.g. various MS Office components
			ContinueLoop
		ElseIf RegRead($sKey, "SystemComponent") Then
			; SystemComponent generally denotes some dependency sort of thing that shouldn't be
			; uninstalled by itself
			ContinueLoop
		EndIf

		$sUninstall = RegRead($sKey, "UninstallString")
		If @error Or $sUninstall = "" Then ContinueLoop

		; if it's a Win Installer issue, make sure the "uninstall string" actually UNINSTALLS
		$sUninstall = StringReplace($sUninstall, "msiexec.exe /i", "msiexec.exe /X")
		; then, add the /passive switch for an unattended but not silent uninstallation
		If StringInStr($sUninstall, "msiexec.exe") Then $sUninstall &= " /passive"

		; read in various standard info
		$sPub = RegRead($sKey, "Publisher")
		$sVer = RegRead($sKey, "DisplayVersion")
		$sLoc = RegRead($sKey, "InstallLocation")
		$sDate = RegRead($sKey, "InstallDate")
		$sIcon = RegRead($sKey, "DisplayIcon")
		If RegRead($sKey, "EstimatedSize") Then
			$iSize = NumberPadZeroesFloat(Round(Number(RegRead($sKey, "EstimatedSize")) / 1024, 2))
			$iSizeTotal += $iSize
		Else
			; this is where manual size calculation would go
			; obviously the first path to use would be that in InstallLocation
			; provided it isn't bullshit like "c:\windows"
			$iSize = ""
		EndIf

		; update progress display
		GUICtrlSetData($lStatus, $i+1 & " / " & UBound($aSubkeys) & ". " & Round((($i+1)/UBound($aSubkeys))*100) & "% done")
		GUICtrlSetData($hProgress, Round((($i+1)/UBound($aSubkeys))*100))

		; create the listview item & stick strings in subitems
		$hItem = GUICtrlCreateListViewItem($sName, $ListView1)
		$iCurrent = _GUICtrlListView_GetItemCount($ListView1) - 1
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sPub, $iColPub)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sVer, $iColVersion)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $iSize, $iColSize)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sDate, $iColDate)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sUninstall, $iColUninstall)
		_GUICtrlListView_SetItemText($ListView1, $iCurrent, $sLoc, $iColLocation)

		; if there's something listed in DisplayIcon, separate the filepath from the resource index
		If $sIcon <> "" Then
			$aIcon = StringSplit($sIcon, ",")
			If $aIcon[0] > 1 Then
				$hIcon = _GUIImageList_AddIcon($hImage, $aIcon[1], $aIcon[2])
			Else
				$hIcon = _GUIImageList_AddIcon($hImage, $aIcon[1], 0)
			EndIf
			_GUICtrlListView_SetItemImage($ListView1, $iCurrent, $hIcon)
		; if the app key is a GUID, it's going to be a Win Installer case, which means the icon is pointed to elsewhere
		ElseIf _stringIsGUID($sEnum) Then
			$sNewKey = "HKLM\Software\Classes\Installer\Products\" & _stringConvertUUID($sEnum, 2)

			$sIcon = RegRead("\\" & $sComp & "\" & $sNewKey, "ProductIcon")
			if NOT @error Then
				; if the icon path exists, use it
				$hIcon = _GUIImageList_AddIcon($hImage, $sIcon)
				_GUICtrlListView_SetItemImage($ListView1, $iCurrent, $hIcon)
			EndIf
		EndIf
	Next

	_GUICtrlListView_SetColumn($ListView1, $iColSize, "Size : " & Round($iSizeTotal), -1, 1)
	_GUICtrlListView_SetColumnWidth($ListView1, $iColName, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($ListView1, $iColPub, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($ListView1, $iColVersion, $LVSCW_AUTOSIZE)
EndFunc   ;==>populateUninstalls

Func getSizes()
	; simple, temporary directory size calculator (in tma2.au3)
	; avoids using DirGetSize so as to not lock everything the fuck up if
	;  the directory happens to contain a horrid mass of files.
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

	; placeholder for future manual GUI bullshit that i hate having to fuck with
	; goddammit

	; return value to let AutoIt go about its GUI business
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

Func WM_COMMAND($hWnd, $iMsg, $iwParam, $ilParam)
	ConsoleWrite("hWnd:    " & $hWnd & @CRLF)
	ConsoleWrite("iMsg:    " & $iMsg & @CRLF)
	ConsoleWrite("iwParam: " & $iwParam & @CRLF)
	ConsoleWrite("ilParam: " & $ilParam & @CRLF)

	; will probably only be needed in case of manual context menu handling
	; e.g. to handle a menu item click

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

Func WM_CONTEXTMENU($hWnd, $iMsg, $iwParam, $ilParam)
	; stupid bullshit manual context menu generation & display
	; i hate it i hate it
	; plus whatever stupid crap i put in this function will probably be re-written
	;  'cause you have to intercept right-clicks on header controls yourself,
	;  otherwise it'll act as a part of the ListView as a whole
	Local $hMenu

	$hMenu = _GUICtrlMenu_CreatePopup()

	For $i = 0 To _GUICtrlListView_GetColumnCount($ListView1) - 1
		$aTemp = _GUICtrlListView_GetColumn($ListView1, $i)
		_GUICtrlMenu_AddMenuItem($hMenu, $aTemp[5], $hMenuItem[$i])
	Next

	_GUICtrlMenu_TrackPopupMenu($hMenu, $iwParam)
	_GUICtrlMenu_DestroyMenu($hMenu)
EndFunc   ;==>WM_CONTEXTMENU
