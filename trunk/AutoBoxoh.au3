Global $debug = 1
Global $bGo = 0
Global $pathHistory = @AppDataDir & "\autoboxoh.txt"

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <TreeViewConstants.au3>
#include <WindowsConstants.au3>

#include <guicombobox.au3>
#include <guitreeview.au3>
#include <guiListView.au3>
#include <_xmldomwrapper.au3>
#include <array.au3>
#include <inet.au3>

Opt("GUIOnEventMode", 1)
Opt("TrayAutoPause", 0)
OnAutoItExitRegister("autoit_exit")

#Region ### START Koda GUI section ### Form=C:\Program Files (x86)\AutoIt3\Extras\Koda\Forms\AutoBoxoh.kxf
$Form1 = GUICreate("AutoBoxoh2", 314, 273, 248, 338, BitOR($WS_MINIMIZEBOX,$WS_SIZEBOX,$WS_THICKFRAME,$WS_SYSMENU,$WS_CAPTION,$WS_POPUP,$WS_POPUPWINDOW,$WS_GROUP,$WS_BORDER,$WS_CLIPSIBLINGS))
$cmbNumber = GUICtrlCreateCombo("", 8, 8, 217, 25)
GUICtrlSetResizing($cmbNumber, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM+$GUI_DOCKHEIGHT)
$btnGet = GUICtrlCreateButton("Track", 232, 8, 75, 25, 0)
GUICtrlSetResizing($btnGet, $GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
$lblMain = GUICtrlCreateLabel("Status:", 10, 40, 294, 24, $SS_CENTER)
GUICtrlSetFont($lblMain, 12, 800, 0, "Century Gothic")
GUICtrlSetColor($lblMain, 0x0066CC)
GUICtrlSetResizing($lblMain, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKHEIGHT)
$tvMain = GUICtrlCreateTreeView(8, 72, 297, 193)
GUICtrlSetResizing($tvMain, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $aSize = WinGetPos($Form1)
Global $iMaxX = $aSize[2]; + 16
Global $iMaxY = $aSize[3]; + 38
GUIRegisterMsg($WM_GETMINMAXINFO, "WM_GETMINMAXINFO")

GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_EVENT_CLOSE")
GUICtrlSetOnEvent($btnGet, "btnGet_Click")


If FileExists($pathHistory) Then
	GUICtrlSetData($cmbNumber, FileRead($pathHistory))
EndIf

While 1
	Sleep(10)
WEnd

Func btnGet_Click()
	; fedex - 12
	; usps - 22 num
	; ups - 18 alnum
	$sTracking = StringStripWS(StringUpper(GUICtrlRead($cmbNumber)), 8)
	If Not isValidTracking($sTracking) Then
		_error("Invalid tracking number.")
		Return
	EndIf

	$sURL = "http://boxoh.com/?rss=1&t=" & $sTracking
	If $debug Then ConsoleWrite($sURL & @CRLF)

	$sResponse = _INetGetSource($sURL)
	If @error Then
		_error("Couldn't access Boxoh.")
		Return
	EndIf
	If $debug Then ConsoleWrite($sResponse & @CRLF)

	_XMLLoadXML($sResponse)

	;Package update on 06/05/10 12:11 am
	;&lt;br/&gt;&lt;b&gt;
	;Electronic Shipping Info Received
	;&lt;/b&gt;&lt;br/&gt;
	;Bothell, WA

	$aDescription = _XMLGetValue("/rss/channel/description")

	if $aDescription[1] = "No record of that item" Then
		_error("Item not found.")
		Return
	EndIf

	$iValues = _XMLGetValue("/rss/channel/item/description")
	If @error Then
		_error("XMLGetValue error: " & @error & " / " & @extended)
		Return
	EndIf

	$h1 = _GUICtrlTreeView_FindItem($tvMain, $sTracking)
	If Not $h1 Then
		$h1 = _GUICtrlTreeView_Add($tvMain, 0, $sTracking)

	EndIf
	_GUICtrlTreeView_DeleteChildren($tvMain, $h1)

	For $i = 1 To $iValues[0]
		$iValues[$i] = StringReplace(StringReplace(StringReplace($iValues[$i], "<br/><b>", "|"), "</b><br/>", "|"), "Package update on ", "")
		$aNew = StringSplit($iValues[$i], "|")
		_GUICtrlTreeView_AddChild($tvMain, $h1, $aNew[1] & ": " & $aNew[2] & " (" & $aNew[3] & ")")
	Next
	_GUICtrlTreeView_Expand($tvMain)
	#cs
		_GUICtrlListView_DeleteAllItems($lMain)
		For $i=1 To $iValues[0]
		_GUICtrlListView_AddItem($lMain, $iValues[$i])
		_GUICtrlListView_SetItemText($lblMain, $i-1, $sTracking, 1)
		Next
		_GUICtrlListView_SetColumnWidth($lMain, 0, $LVSCW_AUTOSIZE)
	#ce
	_upStatus(_Now())
EndFunc   ;==>btnGet_Click

Func isValidTracking($sNum)
	If StringLen($sNum) = 12 Then
		If StringIsAlNum($sNum) Then Return 1
	ElseIf StringLen($sNum) = 22 Then
		If StringIsDigit($sNum) Then Return 1
	ElseIf StringLen($sNum) = 18 Then
		If StringIsAlNum($sNum) Then Return 1
	EndIf

	Return 0
EndFunc   ;==>isValidTracking

Func _upStatus($sMessage)
	GUICtrlSetData($lblMain, "Status: " & $sMessage)
	GUICtrlSetColor($lblMain, 0x0066CC)
EndFunc   ;==>_upStatus

Func _error($sMessage)
	if @Compiled Then
		GUICtrlSetData($lblMain, $sMessage)
		GUICtrlSetColor($lblMain, 0xFF4444)
	Else
		ConsoleWriteError($sMessage & @CRLF)
	EndIf

EndFunc   ;==>_error

Func GUI_EVENT_CLOSE()
	FileDelete(@TempDir & "\AutoBoxoh.xml")

	$hFile = FileOpen($pathHistory, 2)
	$sHistory = _GUICtrlComboBox_GetList(GUICtrlGetHandle($cmbNumber))
	ConsoleWrite($sHistory & @CRLF)
	FileWrite($hFile, $sHistory)
	if @error Then _error("couldn't update history")
	FileClose($hFile)

	GUIDelete()
	Exit
EndFunc   ;==>GUI_EVENT_CLOSE

Func autoit_exit()
EndFunc   ;==>autoit_exit

Func WM_GETMINMAXINFO($hWnd2, $MsgID, $wParam, $lParam)
	#forceref $MsgID, $wParamFunctionsFunctions
	Local $minmaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)

	DllStructSetData($minmaxinfo, 7, $iMaxX); min width
	DllStructSetData($minmaxinfo, 8, $iMaxY); min height

EndFunc   ;==>WM_GETMINMAXINFO
