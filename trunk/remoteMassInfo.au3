; #INDEX# =======================================================================================================================
; Title .........: remoteMassInfo
; AutoIt Version : 3.3.6.1
; Description ...: For use by Mosaic Technology Group / Dell only. No copying or re-use is allowed.
; Company .......: Mosaic Technology
; Author ........: Jon Dunham
; To-do .........: make them pings and queries <semi|a>-synchronous, boooieeee
; ===============================================================================================================================

#AutoIt3Wrapper_UseX64=N
#AutoIt3Wrapper_Res_ProductVersion=0.5.5
#AutoIt3Wrapper_Res_Fileversion=0.5.5.8
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_Description=remoteMassInfo
#AutoIt3Wrapper_Res_LegalCopyright=© Jon Dunham
#AutoIt3Wrapper_Res_Field=Company|Mosaic Technology
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Icon=Icons\TMA2\remoteMassInfo.ico

#AutoIt3Wrapper_Change2CUI=N
;~ #AutoIt3Wrapper_Run_After=copy %out% "\\ftwgen01\this\field services\autoitscripts\remoteMassInfo 0.5 unstable.exe"
;~ #AutoIt3Wrapper_Run_After=copy  "\\ftwgen01\this\field services\autoitscripts\Icons\Gnome\16\

Global $oRefCount = ObjCreate("WbemScripting.SWbemNamedValueSet")
if @error Then
	MsgBox(4096+16, "You Got Problems", _
		"Whichever machine you're running this one is missing a critical scripting component, namely WbemScripting. " & _
		"If you're missing that, chances are you're missing a few other things. I'm going to go ahead and... terminate this for you. " & _
		"Please feel free to re-run it when you're on a computer built in the last 20 years.", 30)
	Exit 1
EndIf

$oRefCount.Add("GlobalAsync", 0)

$oWbemSink = ObjCreate("WbemScripting.SWbemSink")
ObjEvent($oWbemSink, "__WbemSink_")

Opt("GUIOnEventMode", 1)
Opt("GUICloseOnESC", 0)

If Not @Compiled Then
	AutoItSetOption("TrayIconDebug", 1)
EndIf

Global $beta = True
Global $debugLevel = 4 ; 0 - 5, nothing - debug+
Global $logToDB = True
Global $pathRemoteDB = @AppDataDir & "\TMA2 Tools\remoteDB.sqlite"
Global $aDBColumns[

Global $oError = ObjEvent("AutoIt.Error", "AutoItErr")

Global Const $pingError[4] = ["Offline", "Unreachable", "Bad destination", "Unknown (bad name)"]

if $beta Then
	Global $version = FileGetVersion(@ScriptFullPath) & " UNSTABLE"
Else
	Global $version = FileGetVersion(@ScriptFullPath)
EndIf

Global $bStop = False

#region user settings
Global $pathINI = @AppDataDir & "\TMA2 Tools\remoteMassInfo.ini"

; COOL, LEARN MORE COMPUTERING TALKING LIKE THESE
Global Const $SHOWFLAG_HW = 0x01
Global Const $SHOWFLAG_MON = 0x02
Global Const $SHOWFLAG_NET = 0x04
Global Const $SHOWFLAG_APPV = 0x08
Global Const $SHOWFLAG_CCM = 0x10
Global Const $SHOWFLAG_CERT = 0x20


Global $checkAppV = Int(IniRead($pathINI, "Settings", "checkAppV", 0))
Global $checkCert = Int(IniRead($pathINI, "Settings", "checkCert", 0))
Global $checkCCM = Int(IniRead($pathINI, "Settings", "checkCCM", 1))
Global $checkMon = Int(IniRead($pathINI, "Settings", "checkMon", 0))
Global $checkHw = Int(IniRead($pathINI, "Settings", "checkHw", 1))
Global $checkSw = Int(IniRead($pathINI, "Settings", "checkSw", 1))
;~ Global $checkNet = Int(IniRead($pathINI, "Settings", "checkNet", ""))

Dim $initWidth = @DesktopWidth - 400
Dim $initHeight = @DesktopHeight - 500

Global $iColSelected
#endregion

; Standard
#include <constants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <progressConstants.au3>

; UDFs
#include <guiEdit.au3>
#include <guiListView.au3>
#include <guiStatusBar.au3>
#include <guiRebar.au3>
#include <guiToolbar.au3>
#include <guiimagelist.au3>
#include <guiMenu.au3>
#include <guiButton.au3>

#include <string.au3>
#include <array.au3>
#include <date.au3>
#include <excel.au3>

; wee here we go
#include <sqlite.au3>
#include <sqlite.dll.au3>

; Extra
#include <adfunctions.au3>
#include <ModernMenuRaw.au3>

; personal
#include <compInfo.au3>
#include <_monitorInfo.au3>
#include <_services.au3>
#include <tma2.au3>


#region ### START Koda GUI section ### Form=

; Form1 GUI
; ---------------------------------------------------------------

$Form1 = GUICreate("remoteMassInfo " & $version, $initWidth, $initHeight, 50, 100, BitOR($WS_CLIPCHILDREN, $WS_SIZEBOX, $WS_MINIMIZEBOX, $WS_MAXIMIZEBOX ))
$f1Size = WinGetClientSize($Form1)
GUISetOnEvent($GUI_EVENT_CLOSE, "Form1_Close")


$menuFile = GUICtrlCreateMenu("&File")
$hmenuFile = GUICtrlGetHandle($menuFile)
$menuFileCSV = _GUICtrlCreateODMenuItem("Save to .CSV..." & @TAB & "Ctrl+S", $menuFile, "Icons\Gnome2\16\text-x-generic.ico", 0)
$menuFileXLS = _GUICtrlCreateODMenuItem("Create Excel Sheet" & @TAB & "Ctrl+Shift+S", $menuFile, "Icons\Gnome2\16\x-office-spreadsheet.ico", 0)
_GUICtrlCreateODMenuItem("", $menuFile)
$menuFileQuit = _GUICtrlCreateODMenuItem("Quit" & @TAB & "Ctrl-Q", $menuFile, "Gnome2\16\Application-Exit.ico", 0)

$menuSettings = GUICtrlCreateMenu("&Settings")
$hmenuSettings = GUICtrlGetHandle($menuSettings)
$menuSettingsAppV = _GUICtrlCreateODMenuItem("Check for App-V", $menuSettings, "Icons\TMA2\App-AppV-default.ico")
$menuSettingsCCM = _GUICtrlCreateODMenuItem("Check for SCCM", $menuSettings, "Icons\TMA2\App-SCCM.ico")
$menuSettingsCert = _GUICtrlCreateODMenuItem("Check for Machine Certificate", $menuSettings, "Icons\Gnome\16\stock_certificate.ico")
_GUICtrlCreateODMenuItem("", $menuSettings)
$menuSettingsHw = _GUICtrlCreateODMenuItem("Display Hardware Info", $menuSettings, "Icons\Gnome2\16\audio-card.ico")
$menuSettingsSw = _GUICtrlCreateODMenuItem("Display Software Info", $menuSettings, "Icons\Gnome2\16\applications-other.ico")
$menuSettingsMon = _GUICtrlCreateODMenuItem("Display Monitor Info", $menuSettings, "Icons\Gnome2\16\video-display.ico")
_GUICtrlCreateODMenuItem("", $menuSettings)
$menuSettingsCols = _GUICtrlCreateODMenu("Display Columns", $menuSettings)


$menuHelp = GUICtrlCreateMenu("&Help")
$menuHelpCon = _GUICtrlCreateODMenuItem("Show Output Log", $menuHelp, "Icons\Gnome2\16\utilities-terminal.ico")

if $checkAppV Then
	GUICtrlSetState($menuSettingsAppV, $GUI_CHECKED)
Else
	GUICtrlSetState($menuSettingsAppV, $GUI_UNCHECKED)
EndIf
if $checkCCM Then
	GUICtrlSetState($menuSettingsCCM, $GUI_CHECKED)
Else
	GUICtrlSetState($menuSettingsCCM, $GUI_UNCHECKED)
EndIf
if $checkCert Then
	GUICtrlSetState($menuSettingsCert, $GUI_CHECKED)
Else
	GUICtrlSetState($menuSettingsCert, $GUI_UNCHECKED)
EndIf
if $checkHw Then
	GUICtrlSetState($menuSettingsHw, $GUI_CHECKED)
Else
	GUICtrlSetState($menuSettingsHw, $GUI_UNCHECKED)
EndIf
if $checkSw Then
	GUICtrlSetState($menuSettingsSw, $GUI_CHECKED)
Else
	GUICtrlSetState($menuSettingsSw, $GUI_UNCHECKED)
EndIf
if $checkMon Then
	GUICtrlSetState($menuSettingsMon, $GUI_CHECKED)
Else
	GUICtrlSetState($menuSettingsMon, $GUI_UNCHECKED)
EndIf


GUICtrlSetOnEvent($menuSettingsAppV, "menuSettingsAppV_Click")
GUICtrlSetOnEvent($menuSettingsCCM, "menuSettingsCCM_Click")
GUICtrlSetOnEvent($menuSettingsCert, "menuSettingsCert_Click")
GUICtrlSetOnEvent($menuSettingsHw, "menuSettingsHw_Click")
GUICtrlSetOnEvent($menuSettingsSw, "menuSettingsSw_Click")
GUICtrlSetOnEvent($menuSettingsMon, "menuSettingsMon_Click")

;~ GUICtrlSetOnEvent($menuSettings, "menuSettings_Click")

GUICtrlSetOnEvent($menuFileQuit, "menuFileQuit_Click")

GUICtrlSetOnEvent($menuHelpCon, "menuHelpCon_Click")

; Child output
; ---------------------------------------------------------------
Global $hConsoleGUI, $hConsoleEdit, $ctxHeader

_guiCreateSimpleConsole($hConsoleGUI, $hConsoleEdit, "Logging", 669, 333, BitOR($WS_SIZEBOX, $WS_CAPTION, $WS_GROUP, $WS_CLIPSIBLINGS), $WS_EX_TOOLWINDOW, $Form1)
GUISetOnEvent($GUI_EVENT_CLOSE, "hConsoleGUI_Close", $hConsoleGUI)
WinSetTrans($hConsoleGUI, "", 220)

GUISwitch($Form1)

$progress = GUICtrlCreateProgress(0, 0, -1, 25, $PBS_SMOOTH)

#region REBAR

; Rebar
; ---------------------------------------------------------------
$hRebar = _GUICtrlRebar_Create($Form1, BitOR($CCS_TOP, $RBS_VARHEIGHT, $RBS_AUTOSIZE, $CCS_NODIVIDER))


; Toolbar & related controls
; ---------------------------------------------------------------
$hImageTB = _GUIImageList_Create(16, 16, 5)
_GUIImageList_AddIcon($hImageTB, "Icons\Gnome2\16\media-playback-start.ico")
_GUIImageList_AddIcon($hImageTB, "Icons\Gnome2\16\dialog-information.ico")
_GUIImageList_AddIcon($hImageTB, "Icons\Gnome2\16\media-playback-stop.ico")
_GUIImageList_AddIcon($hImageTB, "Icons\TMA2\remoteWmiInfo.ico")
_GUIImageList_AddIcon($hImageTB, "Icons\Gnome\16\preferences-desktop-remote-desktop.ico")
_GUIImageList_AddIcon($hImageTB, "Icons\Gnome2\16\text-x-generic.ico")
_GUIImageList_AddIcon($hImageTB, "Icons\Gnome2\16\x-office-spreadsheet.ico")
_GUIImageList_AddIcon($hImageTB, "Icons\Gnome\16\stock_data-edit-sql-query.ico")

$hToolbar = _GUICtrlToolbar_Create($Form1, BitOR($TBSTYLE_LIST,$CCS_ADJUSTABLE, $TBSTYLE_FLAT, $CCS_NORESIZE, $CCS_NOPARENTALIGN), BitOR($TBSTYLE_EX_MIXEDBUTTONS, $TBSTYLE_EX_DOUBLEBUFFER))
_GUICtrlToolbar_SetImageList($hToolbar, $hImageTB)
;~ _GUICtrlToolbar_SetButtonSize($hToolbar, 20, 20)

$strtbPing = _GUICtrlToolbar_AddString($hToolbar, "Ping")
$strtbQuery = _GUICtrlToolbar_AddString($hToolbar, "Query")
$strtbStop = _GUICtrlToolbar_AddString($hToolbar, "Stop")
$strtbWMI = _GUICtrlToolbar_AddString($hToolbar, "Run rWMII")
$strtbRC = _GUICtrlToolbar_AddString($hToolbar, "LD Remote")
$strtbCSV = _GUICtrlToolbar_AddString($hToolbar, "Export (CSV)")
$strtbXLS = _GUICtrlToolbar_AddString($hToolbar, "Export (XLS)")
$strtbSQL = _GUICtrlToolbar_AddString($hToolbar, "Update DB")

Global $iItem
Global Enum $tbPing = 1000, $tbQuery, $tbStop, $tbWMI, $tbRC, $tbCSV, $tbXLS
_GUICtrlToolbar_AddButton($hToolbar, $tbPing, 0, $strtbPing, $BTNS_AUTOSIZE)
_GUICtrlToolbar_AddButton($hToolbar, $tbQuery, 1, $strtbQuery, $BTNS_AUTOSIZE)
_GUICtrlToolbar_AddButton($hToolbar, $tbStop, 2, $strtbStop, $BTNS_AUTOSIZE)
_GUICtrlToolbar_AddButtonSep($hToolbar)
_GUICtrlToolbar_AddButton($hToolbar, $tbWMI, 3, $strtbWMI, $BTNS_AUTOSIZE)
_GUICtrlToolbar_AddButton($hToolbar, $tbRC, 4, $strtbRC, $BTNS_AUTOSIZE)
_GUICtrlToolbar_AddButtonSep($hToolbar)
_GUICtrlToolbar_AddButton($hToolbar, $tbCSV, 5, $strtbCSV, $BTNS_AUTOSIZE)
_GUICtrlToolbar_AddButton($hToolbar, $tbXLS, 6, $strtbXLS, $BTNS_AUTOSIZE)
_GUICtrlToolbar_AddButton($hToolbar, $tbSQL, 7, $strtbSQL, BitOR($BTNS_AUTOSIZE,$BTNS_CHECK))

; Rebar... bars
; ---------------------------------------------------------------
_GUICtrlRebar_AddToolBarBand($hRebar, $hToolbar, "Toolbar", -1, $RBBS_TOPALIGN)
_GUICtrlRebar_AddBand($hRebar, GUICtrlGetHandle($progress), 50, 200, "Progress")

#endregion

; Status bar
; ---------------------------------------------------------------
$statusBar2 = _GUICtrlStatusBar_Create($Form1)
Dim $StatusBar2_PartsWidth[3] = [120, 250, -1]
_GUICtrlStatusBar_SetParts($statusBar2, $StatusBar2_PartsWidth)
_GUICtrlStatusBar_SetText($statusBar2, @LogonDomain & "\" & @UserName, 0)
;~ _GUICtrlStatusBar_EmbedControl($statusBar2, 1, GUICtrlGetHandle($progress))
_GUICtrlStatusBar_SetText($statusBar2, "", 2)
;~ _GUICtrlStatusBar_SetMinHeight($statusBar2, 20)
GUICtrlSetResizing($statusBar2, $GUI_DOCKSTATEBAR)

#region ListView

; Calc new client area to position listview
; ---------------------------------------------------------------
$f1Size = WinGetClientSize($Form1)
$aMenuInfo = _GUICtrlMenu_GetMenuBarInfo($Form1)
$iMenuH = $aMenuInfo[3]-$aMenuInfo[1]
$iStatusH =  _GUICtrlStatusBar_GetHeight($statusBar2)
$iRebarH = _GUICtrlRebar_GetBarHeight($hRebar)


dbout("form1 client height " & $f1Size[1], 5)
dbout("menubar height " & $iMenuH, 5)
dbout("rebar height " & $iRebarH, 5)
dbout("rebar pos y " & ControlGetPos($Form1, "", $hRebar), 5)
dbout("status height " & $iStatusH, 5)

dbout("remaining client area " & $f1Size[1]-$iMenuH-$iStatusH-$iRebarH, 5)

;~ Global $list1
$list1 = GUICtrlCreateListView("", 0, $iRebarH, $f1Size[0], $f1Size[1]-$iRebarH-$iStatusH-6, BitOR($WS_BORDER, $LVS_REPORT, $LVS_SHOWSELALWAYS), BitOR($LVS_EX_DOUBLEBUFFER, $LVS_EX_SUBITEMIMAGES, $LVS_EX_SNAPTOGRID, $LVS_EX_CHECKBOXES, $LVS_EX_HEADERDRAGDROP, $LVS_EX_FULLROWSELECT))
;~ $list1 = _GUICtrlListView_Create($Form1, "", 0, $iRebarH, $f1Size[0], $f1Size[1]-$iRebarH-$iStatusH-6, BitOR($WS_BORDER, $LVS_REPORT, $LVS_SHOWSELALWAYS), BitOR($LVS_EX_DOUBLEBUFFER, $LVS_EX_SUBITEMIMAGES, $LVS_EX_SNAPTOGRID, $LVS_EX_CHECKBOXES, $LVS_EX_HEADERDRAGDROP, $LVS_EX_FULLROWSELECT))
$hListView = GUICtrlGetHandle($list1)
; Add columns
; ---------------------------------------------------------------
Global $iColName = _GUICtrlListView_AddColumn($list1, "Computer Name")
Global $iColReply = _GUICtrlListView_AddColumn($list1, "Reply")

Global $iColTime = _GUICtrlListView_AddColumn($list1, "Query Time")
Global $iColTS = _GUICtrlListView_AddColumn($list1, "Contact Timestamp")
If @Compiled Then
	; people generally don't need these
	_GUICtrlListView_HideColumn($list1, $iColTime)
	_GUICtrlListView_HideColumn($list1, $iColTS)
EndIf

Global $iColIP = _GUICtrlListView_AddColumn($list1, "IP")
Global $iColMAC = _GUICtrlListView_AddColumn($list1, "MAC")
Global $iColOnDomain = _GUICtrlListView_AddColumn($list1, "On Domain")
Global $iColLocation = _GUICtrlListView_AddColumn($list1, "Location")
Global $iColOU = _GUICtrlListView_AddColumn($list1, "OU")

Global $iColModel = _GUICtrlListView_AddColumn($list1, "Model")
Global $iColSerial = _GUICtrlListView_AddColumn($list1, "Serial")
Global $iColAsset = _GUICtrlListView_AddColumn($list1, "Asset")
Global $iColRAM = _GUICtrlListView_AddColumn($list1, "RAM (GB)")
Global $iColBIOS = _GUICtrlListView_AddColumn($list1, "BIOS Version")
Global $iColBIOSPwr = _GUICtrlListView_AddColumn($list1, "BIOS Power-On Setting")

Global $iColMonitor = _GUICtrlListView_AddColumn($list1, "Monitor Model")
Global $iColMSerial = _GUICtrlListView_AddColumn($list1, "Monitor Serial")
Global $iColMonitor2 = _GUICtrlListView_AddColumn($list1, "Monitor Model 2")
Global $iColMSerial2 = _GUICtrlListView_AddColumn($list1, "Monitor Serial 2")

Global $iColDefaultUser = _GUICtrlListView_AddColumn($list1, "Default User")
Global $iColCurrentUser = _GUICtrlListView_AddColumn($list1, "Current User")
Global $iColPrinter = _GUICtrlListView_AddColumn($list1, "Default Printer")

Global $iColSoftCitrix = _GUICtrlListView_AddColumn($list1, "Citrix")

; CCM
Global $iColCCM = _GUICtrlListView_AddColumn($list1, "SCCM Installed")
; Cert
Global $iColCertSubject = _GUICtrlListView_AddColumn($list1, "CA Subject Name")
Global $iColCertValid = _GUICtrlListView_AddColumn($list1, "Certificate Valid")
; App-V
Global $iColUserInAppV = _GUICtrlListView_AddColumn($list1, "Default User in App-V Group")
Global $iColInvision = _GUICtrlListView_AddColumn($list1, "Invision")
Global $iColAccessAnyware = _GUICtrlListView_AddColumn($list1, "AccessAnyware")
Global $iColStedmans = _GUICtrlListView_AddColumn($list1, "Stedmans")
Global $iColMIView = _GUICtrlListView_AddColumn($list1, "MIView")
Global $iColMicromedex = _GUICtrlListView_AddColumn($list1, "Micromedex")

GUICtrlSetOnEvent($list1, "list1_Event")
GUICtrlSetResizing($list1, $GUI_DOCKBORDERS)

$imgList = _GUIImageList_Create(16, 16, 5)
$iImgBlank = _GUIImageList_AddIcon($imgList, "")
$iImgOnline = _GUIImageList_AddIcon($imgList, @ScriptDir & "\Icons\Gnome\16\network-online.ico")
$iImgOffline = _GUIImageList_AddIcon($imgList, @ScriptDir & "\Icons\Gnome2\16\network-offline.ico")
$iImgError = _GUIImageList_AddIcon($imgList, @ScriptDir & "\Icons\Gnome2\16\network-error.ico")
$lvImgErr = _GUIImageList_AddIcon($imgList, @ScriptDir & "\Icons\Gnome2\16\dialog-error.ico")
$lvImgWarn = _GUIImageList_AddIcon($imgList, @ScriptDir & "\Icons\Gnome2\16\dialog-warning.ico")
$lvImgOK = _GUIImageList_AddIcon($imgList, @ScriptDir & "\Icons\Gnome2\16\emblem-default.ico")

$lvImgCitrix = _GUIImageList_AddIcon($imgList, @ScriptDir & "\Icons\TMA2\App-Citrix.ico")
$lvImgCitrix923 = _GUIImageList_AddIcon($imgList, @ScriptDir & "\Icons\TMA2\App-Citrix923.ico")
$iImgUnknown = _GUIImageList_AddIcon($imgList, @ScriptDir & "\Icons\Gnome2\16\dialog-question.ico")

_GUICtrlListView_SetImageList($list1, $imgList, 1)

$hHeader = _GUICtrlListView_GetHeader($list1)
$hImageHDR = _GUIImageList_Create(16, 16, 5)
$imgHdrAA = _GUIImageList_AddIcon($hImageHDR, "Icons\TMA2\App-AA.ico")
$imgHdrAppV = _GUIImageList_AddIcon($hImageHDR, "Icons\TMA2\App-AppV.ico")
$imgHdrInv = _GUIImageList_AddIcon($hImageHDR, "Icons\TMA2\App-Inv.ico")
$imgHdrCCM = _GUIImageList_AddIcon($hImageHDR, "Icons\TMA2\App-SCCM.ico")
$imgHdrTHR = _GUIImageList_AddIcon($hImageHDR, "Icons\TMA2\AppTHR.ico")
$imgHdrCert = _GUIImageList_AddIcon($hImageHDR, "Icons\Gnome\16\stock_certificate.ico")

_GUICtrlHeader_SetImageList($hHeader, $hImageHDR)

$ret = _GUICtrlHeader_SetItemImage($hHeader, $iColAccessAnyware, 0)
dbout("header setimage: " & $ret, 4)
_GUICtrlHeader_SetItemImage($hHeader, $iColInvision, 2)
_GUICtrlHeader_SetItemImage($hHeader, $iColStedmans, 1)
_GUICtrlHeader_SetItemImage($hHeader, $iColMicromedex, 1)
_GUICtrlHeader_SetItemImage($hHeader, $iColMIView, 1)
_GUICtrlHeader_SetItemImage($hHeader, $iColCCM, 3)

_GUICtrlHeader_SetItemImage($hHeader, $iColCertSubject, 5)
_GUICtrlHeader_SetItemImage($hHeader, $iColCertValid, 5)

#endregion

$ctxList = GUICtrlCreateContextMenu($list1)
$menuAdd = _GUICtrlCreateODMenuItem("Add Computers..." & @TAB & "Alt+A", $ctxList, "Icons\Gnome2\16\list-add.ico", 0)
$menuPaste = _GUICtrlCreateODMenuItem("Paste Computers" & @TAB & "Ctrl+V", $ctxList, "Icons\Gnome2\16\edit-paste.ico", 0)
$menuCopy = _GUICtrlCreateODMenuItem("Copy" & @TAB & "Ctrl+C", $ctxList, "Icons\Gnome2\16\edit-copy.ico", 0)
_GUICtrlCreateODMenuItem("", $ctxList)
$menuSelAll = _GUICtrlCreateODMenuItem("Select All" & @TAB & "Ctrl+A", $ctxList, "Icons\Gnome\16\stock_select-table.ico", 0)
$menuSelNone = _GUICtrlCreateODMenuItem("Select None" & @TAB & "Ctrl+D", $ctxList, "Icons\Gnome\16\stock_select-none.ico", 0)
$menuSelInv = _GUICtrlCreateODMenuItem("Select Inverse" & @TAB & "Ctrl+I", $ctxList, "Icons\Gnome\16\stock_filters-invert.ico", 0)
_GUICtrlCreateODMenuItem("", $ctxList)
$menuCheck = _GUICtrlCreateODMenuItem("Check" & @TAB & "+", $ctxList, "Icons\Gnome\16\stock_form-checkbox.ico", 0)
$menuUncheck = _GUICtrlCreateODMenuItem("Uncheck" & @TAB & "-", $ctxList, "Icons\Gnome\16\stock_form-uncheckbox.ico", 0)
_GUICtrlCreateODMenuItem("", $ctxList)
$menuDel = _GUICtrlCreateODMenuItem("Delete Selected" & @TAB & "Del", $ctxList, "Icons\Gnome2\16\list-remove.ico", 0)
$menuClr = _GUICtrlCreateODMenuItem("Delete All" & @TAB & "Ctrl+Del", $ctxList, "Icons\Gnome2\16\edit-delete.ico", 0)
$menuClean = _GUICtrlCreateODMenuItem("Delete Offline" & @TAB & "Alt+Del", $ctxList, "Icons\Gnome\16\stock_data-delete-table.ico", 0)

GUICtrlSetOnEvent($menuAdd, "menuAdd_Click")
GUICtrlSetOnEvent($menuPaste, "menuPaste_Click")
GUICtrlSetOnEvent($menuCopy, "menuCopy_Click")
; --
GUICtrlSetOnEvent($menuSelAll, "menuSelAll_Click")
GUICtrlSetOnEvent($menuSelNone, "menuSelNone_Click")
GUICtrlSetOnEvent($menuSelInv, "menuSelInv_Click")
; --
GUICtrlSetOnEvent($menuCheck, "menuCheck_Click")
GUICtrlSetOnEvent($menuUncheck, "menuUncheck_Click")
; --
GUICtrlSetOnEvent($menuDel, "menuDel_Click")
GUICtrlSetOnEvent($menuClr, "menuClr_Click")
GUICtrlSetOnEvent($menuClean, "menuClean_Click")
; --
GUICtrlSetOnEvent($menuFileCSV, "menuCSV_Click")
GUICtrlSetOnEvent($menuFileXLS, "menuXLS_Click")

; Form2 GUI
; ---------------------------------------------------------------

$Form2 = GUICreate("Add Computers", 250, 393, 469, 336, BitOR($WS_SIZEBOX, $WS_CAPTION, $WS_GROUP, $WS_CLIPSIBLINGS), BitOR($WS_EX_MDICHILD,$WS_EX_TOOLWINDOW), $Form1)
$f2Size = WinGetClientSize($Form2)

$editComps = GUICtrlCreateEdit("", 0, 58, 250, 295)
GUICtrlSetResizing(-1, $GUI_DOCKBORDERS)

$btnAdd = GUICtrlCreateButton("&Add", $f2Size[0] - (65 * 3) - (8 * 3), $f2Size[1] - 25 - 8, 65, 25, $BS_DEFPUSHBUTTON)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
GUICtrlSetTip(-1, "Add computers to main form.")

$btnQB = GUICtrlCreateButton("&Format", $f2Size[0] - (65 * 2) - (8 * 2), $f2Size[1] - 25 - 8, 65, 25)
GUICtrlSetTip(-1, "Fix return characters copied from QuickBase grid edit.")
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)

$btnClose = GUICtrlCreateButton("&Close", $f2Size[0] - 65 - 8, $f2Size[1] - 25 - 8, 65, 25)
GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM + $GUI_DOCKSIZE)
GUICtrlSetTip(-1, "Close window.")

GUICtrlCreateLabel("Enter computer names to contact on separate lines. You can also paste a list of computer names from Excel or QuickBase.", 8, 8, 206, 40)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKRIGHT + $GUI_DOCKHEIGHT)

$imgListAdd = _GUIImageList_Create(16, 16, 5)
_GUIImageList_AddIcon($imgListAdd, "Icons\Gnome2\16\list-add.ico")
$imgListQB = _GUIImageList_Create(16, 16, 5)
_GUIImageList_AddIcon($imgListQB, "Icons\quickbase.ico")
$imgListClose = _GUIImageList_Create(16, 16, 5)
_GUIImageList_AddIcon($imgListClose, "Icons\Gnome2\16\window-close.ico")

_GUICtrlButton_SetImageList($btnAdd, $imgListAdd)
_GUICtrlButton_SetImageList($btnClose, $imgListClose)
_GUICtrlButton_SetImageList($btnQB, $imgListQB)

GUISetState(@SW_HIDE, $Form2)

; Form2 events
; ---------------------------------------------------------------
GUISetOnEvent($GUI_EVENT_CLOSE, "Form2_Close")
GUICtrlSetOnEvent($btnAdd, "btnAdd_Click")
GUICtrlSetOnEvent($btnClose, "btnClose_Click")
GUICtrlSetOnEvent($btnQB, "btnQB_Click")

; Misc
; ---------------------------------------------------------------
GUIRegisterMsg($WM_GETMINMAXINFO, "WM_GETMINMAXINFO")
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
GUIRegisterMsg($WM_SIZE, "WM_SIZE")
GUIRegisterMsg($WM_CONTEXTMENU, "WM_CONTEXTMENU")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
GUIRegisterMsg($WM_MENUCOMMAND, "WM_MENUCOMMAND")
GUIRegisterMsg($WM_KILLFOCUS, "WM_KILLFOCUS")
GUIRegisterMsg($WM_SETFOCUS, "WM_SETFOCUS")
_GUICtrlListView_RegisterSortCallBack($list1)
;GUIRegisterMsg($SW_MINIMIZE, "SW_MINIMIZE")

#endregion ### END Koda GUI section ###

Dim $arrAccelerators[13][2] = [["^q", $menuFileQuit], ["{NUMPADADD}", $menuCheck], ["{NUMPADSUB}", $menuUncheck], ["^i", $menuSelInv],["^v", $menuPaste],["!a", $menuAdd],["^s", $menuFileCSV],["^+s", $menuFileXLS],["{DELETE}", $menuDel],["^{DELETE}", $menuClr],["!{DELETE}", $menuClean],["^a", $menuSelAll],["^d", $menuSelNone]]
GUISetAccelerators($arrAccelerators, $Form1)

GUISetState(@SW_SHOW, $Form1)

; jon and the terrible, horrible, no-good, very bad hack
_WinAPI_RedrawWindow($hListView)

dbout("form1 hwnd " & $Form1, 5)
dbout("form2 hwnd " & $Form2, 5)
dbout("form3 hwnd " & $hConsoleGUI, 5)


if $beta AND @Compiled Then
	if NOT Number(IniRead($pathINI, "Settings", "betaAccept", 0)) Then
		GUISetState(@SW_DISABLE, $Form1)
		$iq = MsgBox(4096+1+48, "Warning", "This is a development version. There are lots of instabilities and misc bugs." & @CRLF & _
			"Many controls do not function as expected. It's messy." & @CRLF & _
			"If it hangs up for more than a minute or so, kill it in task manager." & @CRLF & @CRLF & _
			"Run? (this message won't show again if you accept)", 0, $Form1)

		if $iq = 1 Then
			IniWrite($pathINI, "Settings", "betaAccept", 1)
			GUISetState(@SW_ENABLE, $Form1)
		Else
			IniWrite($pathINI, "Settings", "betaAccept", 0)
			GUIDelete()
			Exit
		EndIf
	EndIf
EndIf


if @Compiled Then
	if NOT RegRead("HKCR\CAPICOM.Certificate", "") Then
		if FileCopy("\\ftwgen01\this\field services\autoitscripts\misc\capicom.*", "c:\windows\system32\") Then
			$iReturn = ShellExecuteWait("regsvr32", "/s c:\windows\system32\capicom.dll")
			if not @error and $iReturn = 0 Then
				MsgBox(4096, "CAPICOM Interface", "capicom.dll successfully copied and registered. Restarting...", 5, $Form1)
				Run(@ScriptFullPath & ' /AutoIt3ExecuteLine  "ProcessWaitClose(' & @AutoItPID & ', 5)"')
				Exit
			ElseIf $iReturn <> 0 Then
				MsgBox(4096+16, "CAPICOM Interface", "Error registering dll: " & $iReturn)
				GUICtrlSetState($menuSettingsCert, $GUI_UNCHECKED)
				GUICtrlSetState($menuSettingsCert, $GUI_DISABLE)
				$checkCert = 0
			EndIf
		Else
			$iQ = MsgBox(4096+48+4, "Warning", "The computer you're on lacks a necessary activex component to view certificates. The certificate checking information will be unavailable." & @CRLF & @CRLF & _
				"Would you like to visit the download page for it now?")

			if $iQ = 6 Then
				ShellExecute("iexplore.exe", '"http://www.microsoft.com/downloads/details.aspx?FamilyID=860ee43a-a843-462f-abb5-ff88ea5896f6&displaylang=en"')
				MsgBox(4096+64, "Info", "After you've installed it, start > run: " & @CRLF & _
					"regsvr32 C:\Program Files\Microsoft CAPICOM 2.1.0.2 SDK\Lib\X86\capicom.dll" & @CRLF & _
					"and then re-run this application.")
			EndIf

			GUICtrlSetState($menuSettingsCert, $GUI_UNCHECKED)
			GUICtrlSetState($menuSettingsCert, $GUI_DISABLE)
			$checkCert = 0
		EndIf
	EndIf
EndIf

While 1
	Sleep(10)
WEnd

Func WM_SIZE($hWnd, $iMsg, $iwParam, $ilParam)
	_GUICtrlStatusBar_Resize($statusBar2)
	_SendMessage($hRebar, 0x05)

;~ 	DllCall("user32.dll", "int", "SendMessage", "hwnd", $hToolbar, "int", $TB_AUTOSIZE, "int", 0, "int", 0)
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_SIZE

Func WM_MENUCOMMAND($iwParam, $ilParam)
;~ 	#forceref $hWnd, $iMsg, $iwParam, $ilParam


	dbout("WM_MENUCOMMAND", 5)
;~ 	dbout("hWnd   : " & $hWnd)
;~ 	dbout("iMsg   : " & $iMsg)
	dbout("wParam : " & $iwParam, 5)
	dbout("iParam : " & $ilParam, 5)

;~ 	$tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
	Switch $ilParam
		Case $ctxHeader
			if _GUICtrlHeader_GetItemWidth($hHeader, $iwParam) = 0 Then
				_GUICtrlHeader_SetItemWidth($hHeader, _GUICtrlHeader_OrderToIndex($hHeader, $iwParam), $LVSCW_AUTOSIZE)
			Else
				_GUICtrlHeader_SetItemWidth($hHeader, $iwParam, 0)
			EndIf
	endSwitch
EndFunc

Func WM_COMMAND($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg, $iwParam, $ilParam


	dbout("WM_COMMAND", 5)
;~ 	dbout("hWnd   : " & $hWnd)
;~ 	dbout("iMsg   : " & $iMsg)
	dbout("wParam : " & $iwParam, 5)
	dbout("iParam : " & $ilParam, 5)

;~ 	$tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
EndFunc

Func WM_CONTEXTMENU($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg, $iwParam, $ilParam

	Switch $iwParam
		Case $hListView
			$x = _WinAPI_LoWord($ilParam)
			$y = _WinAPI_HiWord($ilParam)
			$struct = DllStructCreate($tagPoint)
			DllStructSetData($struct, "x", $x)
			DllStructSetData($struct, "y", $y)

			if _WinAPI_WindowFromPoint($struct) = $hHeader Then
				$ctxHeader = _GUICtrlMenu_CreatePopup(8)
				$aCols = _GUICtrlListView_GetColumnOrderArray($list1)
				For $y=1 To $aCols[0]
					$aTemp = _GUICtrlListView_GetColumn($list1, $aCols[$y])
					$iM = _GUICtrlMenu_AddMenuItem($ctxHeader, $aTemp[5])
					if $aTemp[4] = 0 Then
						_GUICtrlMenu_SetItemChecked($ctxHeader, $iM, False)
					Else
						_GUICtrlMenu_SetItemChecked($ctxHeader, $iM)
					EndIf
				Next
				_GUICtrlMenu_TrackPopupMenu($ctxHeader, $hListView)
				_GUICtrlMenu_DestroyMenu($ctxHeader)

			EndIf
	EndSwitch


	dbout("WM_CONTEXTMENU - " & $hWnd, 5)
	dbout("  iwParam " & $iwParam, 5)
	dbout("  iwParam " & $ilParam, 5)

	Return $GUI_RUNDEFMSG
EndFunc

Func HEADERCONTEXT_HAHAHA()
	$ctxHeader = _GUICtrlMenu_CreatePopup(41)
	$aCols = _GUICtrlListView_GetColumnOrderArray($list1)

	For $y=1 To $aCols[0]
		$aTemp = _GUICtrlListView_GetColumn($list1, $aCols[$y])
		$iM = _GUICtrlMenu_AddMenuItem($ctxHeader, $aTemp[5])
		if $aTemp[4] = 0 Then
			_GUICtrlMenu_SetItemChecked($ctxHeader, $iM, False)
		Else
			_GUICtrlMenu_SetItemChecked($ctxHeader, $iM)
		EndIf
	Next

	_GUICtrlMenu_TrackPopupMenu($ctxHeader, $hHeader)
	_GUICtrlMenu_DestroyMenu($ctxHeader)
EndFunc

Func WM_SETFOCUS($iwParam, $ilParam)
	dbout(" -- set focus" & @TAB & $iwParam & "/" & $ilParam, 5)
EndFunc
Func WM_KILLFOCUS($iwParam, $ilParam)
	dbout(" -- kill focus" & @TAB & $iwParam & "/" & $ilParam, 5)
EndFunc

; for toolbar actions
Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)
	#forceref $hWnd, $iMsg, $iwParam, $ilParam
	Local $hwndFrom, $code, $index
	Local $tNMTOOLBAR, $tInfo, $iID

	$tNMTOOLBAR = DllStructCreate($tagNMTOOLBAR, $ilParam)
;~     $tInfo = DllStructCreate($tagNMTTDISPINFO, $ilParam)

	$hwndFrom = DllStructGetData($tNMTOOLBAR, "hWndFrom")
	$code = DllStructGetData($tNMTOOLBAR, "Code")
	$idFrom = DllStructGetData($tNMTOOLBAR, "idFrom")
	$iItem = DllStructGetData($tNMTOOLBAR, "iItem")

	if $hWnd = $hConsoleGUI Then
		if $code = $NM_CUSTOMDRAW Then
			dbout("  customdraw", 5)
			dbout("hWnd: " & $hWnd, 5)
		EndIf
	EndIf

	Switch $hwndFrom
		; toolbar
		Case $hRebar
			Switch $code
				Case $RBN_HEIGHTCHANGE
;~ 					$rebarHeight = _GUICtrlRebar_GetBarHeight($hRebar)
;~ 					$form1size = WinGetClientSize($Form1)

;~ 					GUICtrlSetPos($list1, 0, $rebarHeight+$iMenuH, $form1size[0], $form1size[1]-$rebarHeight-$iStatusH-$iMenuH-6)
			EndSwitch
		Case $hToolbar
			Switch $code
				; http://msdn.microsoft.com/en-us/library/bb760490(v=VS.85).aspx
				Case $NM_CLICK
					$index = _GUICtrlToolbar_CommandToIndex($hToolbar, $iItem)

					If _GUICtrlToolbar_IsButtonEnabled($hToolbar, $iItem) Then
						_GUICtrlToolbar_EnableButton($hToolbar, $tbPing, False)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbQuery, False)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbWMI, False)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbRC, False)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbCSV, False)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbXLS, False)
;~ 						GUICtrlSetState($list1, $GUI_DISABLE)

						Switch $iItem
							Case $tbPing
								tbPing_Click()
							Case $tbQuery
								tbQuery_Click()
							Case $tbStop
								$bStop = True
							Case $tbWMI
								tbWMI_Click()
							Case $tbRC
								$i = _GUICtrlListView_GetSelectionMark($hListView)
								if $i <> -1 Then
									tbRC_Click(_GUICtrlListView_GetItemText($hListView, $i, $iColName))
								EndIf
							Case $tbCSV
								menuCSV_Click()
							Case $tbXLS
								menuXLS_Click()
						EndSwitch

;~ 						GUICtrlSetState($list1, $GUI_ENABLE)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbPing)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbQuery)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbWMI)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbRC)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbCSV)
						_GUICtrlToolbar_EnableButton($hToolbar, $tbXLS)
					EndIf
			EndSwitch
		; listview
		Case $hListView
			$tNMLISTVIEW = DllStructCreate($tagNMLISTVIEW, $ilParam)
			$iItem = DllStructGetData($tNMLISTVIEW, "Item")
			$iSub = DllStructGetData($tNMLISTVIEW, "SubItem")
			Switch $code
				Case $NM_RCLICK
					dbout("listview right-clicked", 5)
					dbout("idFrom " & $idFrom, 5)
					dbout("Item " & $iItem, 5)
					dbout("Subitem " & $iSub, 5)

					$count = _GUICtrlListView_GetSelectedCount($hListView)
					if $count = 1 Then
						_GUICtrlODMenuItemSetText($menuCopy, "Copy " & _GUICtrlListView_GetItemText($hListView, $iItem, $iSub) & @TAB & "Ctrl+C")
					Else
						_GUICtrlODMenuItemSetText($menuCopy, "Copy " & $count & " Items" & @TAB & "Ctrl+C")
					EndIf

					$iColSelected = $iSub
;~ 				Case $NM_CLICK
;~ 					$tInfo = DllStructCreate($tagNMITEMACTIVATE, $ilParam)
;~ 					ConsoleWrite("LISTVIEW SINGLE CLICKED ITEM: " & DllStructGetData($tInfo, "Index") & "  SUBITEM " & DllStructGetData($tInfo, "SubItem") & @CRLF)
			EndSwitch
		Case $hHeader
			$tNMHEADER = DllStructCreate($tagNMHEADER, $ilParam)
			$Item = DllStructGetData($tNMHEADER, "Item")
			Switch $code
				Case $NM_RCLICK

					dbout("header right-clicked", 5)

					Return False
				Case $NM_CLICK
;~ 					_GUICtrlListView_SortItems($list1, $Item)
				Case $NM_DBLCLK
					if DllStructGetData($tNMHEADER, "Button") = 0 Then
						_GUICtrlListView_SetColumnWidth($list1, $Item, $LVSCW_AUTOSIZE)
					EndIf
			EndSwitch
	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

Func Form1_Close()
	IniWrite($pathINI, "Settings", "checkAppV", Int($checkAppV))
	IniWrite($pathINI, "Settings", "checkCCM", Int($checkCCM))
	IniWrite($pathINI, "Settings", "checkCert", Int($checkCert))
	IniWrite($pathINI, "Settings", "checkMon", Int($checkMon))
	IniWrite($pathINI, "Settings", "checkHw", Int($checkHw))
	IniWrite($pathINI, "Settings", "checkSw", Int($checkSw))

	_GUICtrlListView_UnRegisterSortCallBack($list1)
	GUIDelete()
	Exit
EndFunc   ;==>Form1_Close

Func Form2_Close()
	GUISetState(@SW_HIDE, $Form2)
	WinActivate($Form1)
EndFunc   ;==>Form2_Close

Func hConsoleGUI_Close()
	GUISetState(@SW_HIDE, $hConsoleGUI)
	WinActivate($Form1)
EndFunc

Func menuFileQuit_Click()
	Form1_Close()
EndFunc   ;==>menuFileQuit_Click

Func menuHelpCon_Click()
	GUISetState(@SW_SHOW, $hConsoleGUI)
EndFunc


Func list1_Event()
	$tempCol = GUICtrlGetState($list1)
	_GUICtrlListView_SortItems($list1, $tempCol)
EndFunc   ;==>list1_Event

Func menuAdd_Click()
	GUISetState(@SW_SHOW, $Form2)
	GUICtrlSetData($editComps, StringAddCR(ClipGet()))
EndFunc   ;==>menuAdd_Click

#region ------------------------------ Context Menu

Func menuRemote_Click()

EndFunc

Func menuCopy_Click()
	if _GUICtrlListView_GetSelectedCount($hListView) = 0 Then Return

	Local $aItems = _GUICtrlListView_GetSelectedIndices($hListView, True)
;~ 	Local $iCol = _GUICtrlListView_GetSelectedColumn($hListView)
;~ 	dbout("icol " & $iCol)
	Local $sOut

	For $i=1 To $aItems[0]
		$sOut &= _GUICtrlListView_GetItemText($hListView, $aItems[$i], $iColSelected) & @CRLF
	Next

	ClipPut($sOut)
EndFunc

Func menuSelAll_Click()
	For $i = 0 To _GUICtrlListView_GetItemCount($list1) - 1
		_GUICtrlListView_SetItemSelected($list1, $i)
	Next
EndFunc   ;==>menuSelAll_Click

Func menuSelNone_Click()
	For $i = 0 To _GUICtrlListView_GetItemCount($list1) - 1
		_GUICtrlListView_SetItemSelected($list1, $i, False)
	Next
EndFunc   ;==>menuSelNone_Click

Func menuSelInv_Click()
	_GUICtrlListView_BeginUpdate($list1)

	For $i = 0 To _GUICtrlListView_GetItemCount($list1) - 1
		If _GUICtrlListView_GetItemSelected($list1, $i) Then
			_GUICtrlListView_SetItemSelected($list1, $i, False)
		Else
			_GUICtrlListView_SetItemSelected($list1, $i)
		EndIf
	Next

	_GUICtrlListView_EndUpdate($list1)
EndFunc   ;==>menuSelInv_Click

Func menuDel_Click()
	$test = _GUICtrlListView_DeleteItemsSelected($list1)
EndFunc   ;==>menuDel_Click

Func menuClr_Click()
	$test = _GUICtrlListView_DeleteAllItems($list1)
EndFunc   ;==>menuClr_Click

Func menuClean_Click()
	_GUICtrlListView_BeginUpdate($list1)
	Dim $i = 0
	Dim $total = _GUICtrlListView_GetItemCount($list1)

	Do
		If _GUICtrlListView_GetItemImage($list1, $i) = $iImgError Or _GUICtrlListView_GetItemImage($list1, $i) = $iImgOffline Then
			_GUICtrlListView_DeleteItem($list1, $i)
		Else
			$i += 1
		EndIf
	Until $i >= $total
	_GUICtrlListView_EndUpdate($list1)
EndFunc   ;==>menuClean_Click

Func menuCheck_Click()
	__setSelectionChecked()
EndFunc   ;==>menuCheck_Click

Func menuUncheck_Click()
	__setSelectionChecked(False)
EndFunc   ;==>menuUncheck_Click

Func __setSelectionChecked($bCheck = True)
	_GUICtrlListView_BeginUpdate($list1)
	$aSelected = _GUICtrlListView_GetSelectedIndices($list1, True)
	For $i = 1 To $aSelected[0]
		_GUICtrlListView_SetItemChecked($list1, $aSelected[$i], $bCheck)
	Next
	_GUICtrlListView_EndUpdate($list1)
EndFunc

; save to csv
Func menuCSV_Click()
	$pathSave = FileSaveDialog("Select file to save to", _
			@MyDocumentsDir, _
			"Comma-Separated Value (*.csv)", _
			16, _
			"remoteMassInfo-" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & ".csv", _
			$Form1)
	If @error Then Return

	$hFile = FileOpen($pathSave, 2)
	$iColCount = _GUICtrlListView_GetColumnCount($list1)
	; write header
	For $iCol = 0 To $iColCount - 1
		$aColumn = _GUICtrlListView_GetColumn($list1, $iCol)
		If $iCol < $iColCount - 1 Then
			FileWrite($hFile, $aColumn[5] & ",")
		Else
			FileWrite($hFile, $aColumn[5] & @CRLF)
		EndIf
	Next

	; write column data for each row
	For $iItem = 0 To _GUICtrlListView_GetItemCount($list1) - 1
		For $iCol = 0 To $iColCount - 1
			If $iCol < $iColCount - 1 Then
				FileWrite($hFile, _GUICtrlListView_GetItemText($list1, $iItem, $iCol) & ",")
			Else
				FileWrite($hFile, _GUICtrlListView_GetItemText($list1, $iItem, $iCol) & @CRLF)
			EndIf
		Next
	Next
	FileClose($hFile)
	_GUICtrlStatusBar_SetText($statusBar2, _NowTime(5) & " >", 0)
	_GUICtrlStatusBar_SetText($statusBar2, "Saved: " & @MyDocumentsDir & "\remoteMassInfo-" & @YEAR & @MON & @MDAY & "-" & @HOUR & @MIN & @SEC & ".csv", 1)
EndFunc   ;==>menuCSV_Click

; create excel sheet
Func menuXLS_Click()
	$iCount = _GUICtrlListView_GetItemCount($hListView)
	If $iCount = 0 Then
		upStatusBar("No entries to write.")
		Return
	EndIf

	$oExcel = _ExcelBookNew()
	If Not IsObj($oExcel) Then
		MsgBox(0, "rMI", "Could not create object. Excel must be installed first.")
		Return
	Else
		_ExcelSheetNameSet($oExcel, "remoteMassInfo")
	EndIf

	$aCols = _GUICtrlHeader_GetOrderArray($hHeader)

	For $x = 1 To $aCols[0]
		$sText = _GUICtrlHeader_GetItemText($hHeader, $aCols[$x])
		dbout("column " & $aCols[$x] & " text: " & $sText, 4)

		_ExcelWriteCell($oExcel, $sText, 1, $x)
		; set column width somehow?
	Next

	With $oExcel.ActiveSheet
		.Range(.Cells(1, 1), .Cells(1, $aCols[0])).Select
	EndWith

	$oExcel.ActiveWindow.FreezePanes = True
	_ExcelFontSetProperties($oExcel, 1, 1, 1, $aCols[0], True, False, False, 0xFFFFFF, 0x000000)

	; loop through rows
	For $x = 0 To $iCount-1
		for $y=1 to $aCols[0]
			_ExcelWriteCell($oExcel, _GUICtrlListView_GetItemText($hListView, $x, $aCols[$y]), $x+2, $y)
		Next

;~ 		_ExcelWriteArray($oExcel, $x+2, 1, $aItem, 0, 1)
	Next

	$oExcel.Visible = 1

	$oExcel = 0
EndFunc   ;==>menuXLS_Click

Func menuPaste_Click()
	$s = __repairQBLines(ClipGet())

	$aLines = StringSplit($s, @CRLF)

	__addNamesGeneral($aLines)
EndFunc   ;==>menuPaste_Click

#endregion ------------------------------ Context Menu



Func menuSettingsAppV_Click()
	if BitAND(GUICtrlRead($menuSettingsAppV), $GUI_CHECKED) Then
		GUICtrlSetState($menuSettingsAppV, $GUI_UNCHECKED)
		$checkAppV = 0
		_GUICtrlListView_HideColumn($hListView, $iColAccessAnyware)
		_GUICtrlListView_HideColumn($hListView, $iColInvision)
		_GUICtrlListView_HideColumn($hListView, $iColMicromedex)
		_GUICtrlListView_HideColumn($hListView, $iColMIView)
		_GUICtrlListView_HideColumn($hListView, $iColStedmans)
	Else
		GUICtrlSetState($menuSettingsAppV, $GUI_CHECKED)
		$checkAppV = 1
		_GUICtrlListView_SetColumnWidth($hListView, $iColAccessAnyware, $LVSCW_AUTOSIZE)
		_GUICtrlListView_SetColumnWidth($hListView, $iColInvision, $LVSCW_AUTOSIZE)
		_GUICtrlListView_SetColumnWidth($hListView, $iColMicromedex, $LVSCW_AUTOSIZE)
		_GUICtrlListView_SetColumnWidth($hListView, $iColMIView, $LVSCW_AUTOSIZE)
		_GUICtrlListView_SetColumnWidth($hListView, $iColStedmans, $LVSCW_AUTOSIZE)
	EndIf
EndFunc

Func menuSettingsCCM_Click()
	if BitAND(GUICtrlRead($menuSettingsCCM), $GUI_CHECKED) Then
		GUICtrlSetState($menuSettingsCCM, $GUI_UNCHECKED)
		$checkCCM = 0
		_GUICtrlListView_HideColumn($hListView, $iColCCM)
	Else
		GUICtrlSetState($menuSettingsCCM, $GUI_CHECKED)
		$checkCCM = 1
		_GUICtrlListView_SetColumnWidth($hListView, $iColCCM, 100)
	EndIf
EndFunc

Func menuSettingsCert_Click()
	if BitAND(GUICtrlRead($menuSettingsCert), $GUI_CHECKED) Then
		GUICtrlSetState($menuSettingsCert, $GUI_UNCHECKED)
		$checkCert = 0
		_GUICtrlListView_HideColumn($hListView, $iColCertSubject)
		_GUICtrlListView_HideColumn($hListView, $iColCertValid)
	Else
		GUICtrlSetState($menuSettingsCert, $GUI_CHECKED)
		$checkCert = 1
		_GUICtrlListView_SetColumnWidth($hListView, $iColCertSubject, 150)
		_GUICtrlListView_SetColumnWidth($hListView, $iColCertValid, 50)
	EndIf
EndFunc

Func menuSettingsHw_Click()
	if BitAND(GUICtrlRead($menuSettingsHw), $GUI_CHECKED) Then
		GUICtrlSetState($menuSettingsHw, $GUI_UNCHECKED)
		$checkHw = 0
	Else
		GUICtrlSetState($menuSettingsHw, $GUI_CHECKED)
		$checkHw = 1
	EndIf
EndFunc

Func menuSettingsSw_Click()
	if BitAND(GUICtrlRead($menuSettingsSw), $GUI_CHECKED) Then
		GUICtrlSetState($menuSettingsSw, $GUI_UNCHECKED)
		$checkSw = 0
		_GUICtrlListView_HideColumn($hListView, $iColSoftCitrix)
	Else
		GUICtrlSetState($menuSettingsSw, $GUI_CHECKED)
		$checkSw = 1
		_GUICtrlListView_SetColumnWidth($hListView, $iColSoftCitrix, 100)
	EndIf
EndFunc

Func menuSettingsMon_Click()
	if BitAND(GUICtrlRead($menuSettingsMon), $GUI_CHECKED) Then
		GUICtrlSetState($menuSettingsMon, $GUI_UNCHECKED)
		$checkMon = 0
		_GUICtrlListView_HideColumn($hListView, $iColMonitor)
		_GUICtrlListView_HideColumn($hListView, $iColMSerial)
	Else
		GUICtrlSetState($menuSettingsMon, $GUI_CHECKED)
		$checkMon = 1
		_GUICtrlListView_SetColumnWidth($hListView, $iColMonitor, $LVSCW_AUTOSIZE)
		_GUICtrlListView_SetColumnWidth($hListView, $iColMSerial, $LVSCW_AUTOSIZE)
	EndIf
EndFunc

#cs
Func btnProcess_Click()
	$sProcess = GUICtrlRead($inProcess)
	_GUICtrlListView_SetColumn($list1, $iColProcess, $sProcess)

	For $i = 0 To _GUICtrlListView_GetItemCount($list1)
		If _GUICtrlListView_GetItemChecked($list1, $i) = False Or _GUICtrlListView_GetItemGroupID($list1, $i) <> 0 Then
			ContinueLoop
		EndIf

		$sComp = _GUICtrlListView_GetItemText($list1, $i, $iColName)
		$iPID = _processExists($sProcess, $sComp)

		If $iPID Then
			_GUICtrlListView_SetItemText($list1, $i, "Yes (" & $iPID & ")", $iColProcess)
		EndIf
	Next
EndFunc   ;==>btnProcess_Click
#ce

#region ------------------------------ Add form buttons

Func btnAdd_Click()
	$lines = _GUICtrlEdit_GetLineCount($editComps)
	If $lines = 0 Then Return

	Local $a[$lines + 1]
	$a[0] = $lines

	For $i = 1 To $lines
		$a[$i] = _GUICtrlEdit_GetLine($editComps, $i - 1)
	Next

	__addNamesGeneral($a)
EndFunc   ;==>btnAdd_Click

Func __addNamesGeneral($aNames)
	For $i = 1 To $aNames[0]
		If $aNames[$i] Then
			$iRet = GUICtrlCreateListViewItem($aNames[$i], $list1)
;~ 			$index = _GUICtrlListView_MapIDToIndex($list1, $iRet)
;~ 			_GUICtrlListView_SetItemGroupID($list1, $index, 5)
;~ 			dbout("index: " & $index)
		EndIf
	Next

	_GUICtrlListView_SetItemChecked($list1, -1)
	_GUICtrlListView_SetColumnWidth($list1, 0, $LVSCW_AUTOSIZE)
EndFunc   ;==>__addNamesGeneral

Func __repairQBLines($s)
	Return StringAddCR($s)
EndFunc   ;==>__repairQBLines

Func btnQB_Click()
	$newFormat = __repairQBLines(GUICtrlRead($editComps))

	GUICtrlSetData($editComps, $newFormat)
EndFunc   ;==>btnQB_Click

Func btnClose_Click()
	Form2_Close()
EndFunc   ;==>btnClose_Click

#endregion ------------------------------ Add form buttons

#region ------------------------------ Toolbar items

Func tbPing_Click()
	$iCount = _GUICtrlListView_GetItemCount($hListView)
	if $iCount < 1 Then Return
;~ 	TCPStartup()
	Local $iLimit = 20, $sSQLiteDll
	Local $timer = TimerInit()

	if $logToDB Then
		$sSQLiteDll = _SQLite_Startup()
		if @error Then
			$logToDB = False
		Else
			_SQLite_Open($pathRemoteDB)
			_SQLite_Exec("CREATE TABLE IF NOT EXISTS main(
		EndIf
	EndIf

	; ===== DEBUG, CHECK LOAD
	$memStats = ProcessGetStats(-1, 0)
	$ioStats = ProcessGetStats(-1, 1)
	dbout(" --- Process Memory", 3)
	dbout("Memory use      :" & $memStats[0], 3)
	dbout("Peak memory use :" & $memStats[1], 3)
	dbout(" --- Process I/O", 3)
	dbout("Read/Write Ct   :" & $ioStats[0] & "/" & $ioStats[1], 3)
	dbout("Read/Write Bts  :" & $ioStats[3] & "/" & $ioStats[4], 3)
	; ===== >>

	For $i = 0 To _GUICtrlListView_GetItemCount($list1)
		If Not _GUICtrlListView_GetItemChecked($list1, $i) Then
			ContinueLoop
		EndIf

		$compName = _GUICtrlListView_GetItemText($list1, $i)
		_GUICtrlListView_EnsureVisible($hListView, $i)
		_GUICtrlListView_SetItemImage($hListView, $i, 5)

		$percentDone = Round(($i / _GUICtrlListView_GetItemCount($list1)) * 100, 0)
		upStatusBar($i & " / " & _GUICtrlListView_GetItemCount($list1) & " [" & $percentDone & "%], Current: " & $compName)
		WinSetTitle($Form1, "", "[" & $percentDone & "% done] remoteMassInfo " & $version)
		_GUICtrlRebar_SetBandText($hRebar, 1, $percentDone & "%")
		GUICtrlSetData($progress, $percentDone)

		#cs
		if _ADObjectExists($compName, "cn") Then
			_GUICtrlListView_SetItemText($hListView, $i, "Yes", $iColOnDomain)
		Else
			_GUICtrlListView_SetItemText($hListView, $i, "No", $iColOnDomain)
		EndIf
		#ce

		$oPing = _superPing($compName, $oWbemSink, 200, $i)
		$oRefCount.GlobalAsync = $oRefCount.GlobalAsync + 1
;~ 		$oRefCount.Add($compName, $oPing)

		if $oRefCount.GlobalAsync >= $iLimit Then
			dbout("iteration " & $i & ": " & $oRefCount.GlobalAsync & " operations active, waiting for count to decrease...", 3)
			Do
				Sleep(50)
				if $bStop Then
					$bStop = False
					ExitLoop 2
				EndIf
			Until $oRefCount.GlobalAsync < $iLimit/3
		EndIf

		If $bStop Then
			$bStop = False
			ExitLoop
		EndIf
	Next

	; ===== DEBUG, CHECK LOAD
	$memStats = ProcessGetStats(-1, 0)
	$ioStats = ProcessGetStats(-1, 1)
	dbout(" --- Process Memory", 3)
	dbout("Memory use      :" & $memStats[0], 3)
	dbout("Peak memory use :" & $memStats[1], 3)
	dbout(" --- Process I/O", 3)
	dbout("Read/Write Ct   :" & $ioStats[0] & "/" & $ioStats[1], 3)
	dbout("Read/Write Bts  :" & $ioStats[3] & "/" & $ioStats[4], 3)
	; ===== >>

	_GUICtrlListView_SetColumnWidth($list1, $iColName, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($list1, $iColIP, $LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($list1, $iColLocation, $LVSCW_AUTOSIZE)
	if $oRefCount.GlobalAsync > 0 Then
		upStatusBar($i-1 & " computers contacted in " & timeFormat(TimerDiff($timer)) & ". " & $oRefCount.GlobalAsync & " replies still queued.")
	Else
		upStatusBar($i-1 & " computers contacted in " & timeFormat(TimerDiff($timer)) & ".")
	EndIf

	WinSetTitle($Form1, "", "remoteMassInfo " & $version)
;~ 	TCPShutdown()
EndFunc   ;==>tbPing_Click

Func __WbemSink_OnObjectReady($oObject, $oAsyncContext)
	; use "Caller" to identify all script calls
	if $oAsyncContext.Caller.Value = "PingStatus" Then
;~ 		$oPingProperties = $oObject.Properties_
		$oAsyncContext.Add("Reply", $oObject.ResponseTime)
		$oAsyncContext.Add("ResolvedIP", $oObject.ProtocolAddress)

		if VarGetType($oObject.StatusCode) = "String" Then
			$oAsyncContext.PingStatus.Value = "Not found"
			_GUICtrlListView_SetItemText($list1, $oAsyncContext.lParam.Value, $oAsyncContext.PingStatus.Value, $iColReply)
			_GUICtrlListView_SetItemImage($list1, $oAsyncContext.lParam.Value, $iImgUnknown)
			_GUICtrlListView_SetItemChecked($list1, $oAsyncContext.lParam.Value, False)
		ElseIf $oObject.StatusCode = 0 Then
			$oAsyncContext.PingStatus.Value = "Online"
			_GUICtrlListView_SetItemImage($list1, $oAsyncContext.lParam.Value, $iImgOnline)
			_GUICtrlListView_SetItemText($list1, $oAsyncContext.lParam.Value, $oObject.ResponseTime & "ms", $iColReply)
			_GUICtrlListView_SetItemText($list1, $oAsyncContext.lParam.Value, $oObject.ProtocolAddress, $iColIP)
			_GUICtrlListView_SetItemText($list1, $oAsyncContext.lParam.Value, _ipGetLocation($oObject.ProtocolAddress), $iColLocation)
			_GUICtrlListView_SetItemChecked($list1, $oAsyncContext.lParam.Value)
		ElseIf $oObject.StatusCode = 11010 Then
			$oAsyncContext.PingStatus.Value = "Offline"
			_GUICtrlListView_SetItemChecked($list1, $oAsyncContext.lParam.Value, False)
			_GUICtrlListView_SetItemImage($list1, $oAsyncContext.lParam.Value, $iImgError)
			_GUICtrlListView_SetItemText($list1, $oAsyncContext.lParam.Value, $oAsyncContext.PingStatus.Value, $iColReply)
		EndIf

		$oAsyncContext.Remove("Caller")
		$oRefCount.Add($oAsyncContext.Address.Value, $oAsyncContext)
		$oRefCount.GlobalAsync.Value = $oRefCount.GlobalAsync.Value - 1
	ElseIf $oAsyncContext.Caller.Value = "Test" Then
		ConsoleWrite("!Test object ready" & @CRLF)
	EndIf
EndFunc

Func tbRC_Click($sComp)
	$sRCPath = "C:\Program Files\LANDesk\ServerManager\RCViewer\isscntr.exe"
	If FileExists($sRCPath) Then
		;isscntr /a<address> /c<command> /l /s<core server>
		$sCore = "/sftwldmcor01"
		$sAddress = "/a" & $sComp
		$sCommand = '/c"remote control"'

		$sArgs = $sAddress & " " & $sCommand & " " & $sCore

		$iReturn = ShellExecute($sRCPath, $sArgs)
	Else
		Dim $oIE = _IECreate("http://landesk/RemoteSession.aspx?machine=" & $sComp & "&operation=rc", 0, 0)
		_IEQuit($oIE)
	EndIf
EndFunc   ;==>_LDstartRC

; given the AMH set, ~479 computers were queried in ~5:11 using the current method, getting in total the following items:
;		Computer cert from the registry
Func tbQuery_Click()
	$iCount = _GUICtrlListView_GetItemCount($list1)
	if $iCount < 1 Then Return
	Dim $info, $mInfo, $compName, $Contacted, $timerTotal = TimerInit()

	For $i = 0 To _GUICtrlListView_GetItemCount($list1)
		If Not _GUICtrlListView_GetItemChecked($list1, $i) Then
			ContinueLoop
		ElseIf NOT StringRegExp(_GUICtrlListView_GetItemText($hListView, $i, $iColReply), "\d{1,3}ms") Then
			_GUICtrlListView_SetItemChecked($hListView, $i, False)
			ContinueLoop
		EndIf

		$compName = _GUICtrlListView_GetItemText($list1, $i, $iColName)
;~ 		$Contacted = _GUICtrlListView_GetItemText($list1, $i, 4)
		$percentDone = Round(($i / _GUICtrlListView_GetItemCount($list1)) * 100, 0)
		_GUICtrlRebar_SetBandText($hRebar, 1, $percentDone & "%")
		GUICtrlSetData($progress, $percentDone)

		upStatusBar($i & " / " & _GUICtrlListView_GetItemCount($list1) & " [" & $percentDone & "%], Current: " & $compName)
		WinSetTitle($Form1, "", "[" & $percentDone & "% done] remoteMassInfo " & $version)

		$timer = TimerInit()

		if $checkHw Then
			$info = queryComp($compName)
			If @error Then ContinueLoop

			_GUICtrlListView_SetItemText($list1, $i, $info[0], $iColModel)
			_GUICtrlListView_SetItemText($list1, $i, $info[1], $iColSerial)
			_GUICtrlListView_SetItemText($list1, $i, $info[2], $iColAsset)
			_GUICtrlListView_SetItemText($list1, $i, $info[3], $iColBIOS)
			_GUICtrlListView_SetItemText($list1, $i, $info[4] & " GB", $iColRAM)
			_GUICtrlListView_SetItemText($list1, $i, $info[5], $iColDefaultUser)
			_GUICtrlListView_SetItemText($list1, $i, $info[7], $iColCurrentUser)
			_GUICtrlListView_SetItemText($list1, $i, $info[6], $iColOU)
			_GUICtrlListView_SetItemText($list1, $i, $info[8], $iColPrinter)
			_GUICtrlListView_SetItemText($list1, $i, $info[9], $iColMAC)

			; too slow, fix this bullshit
			#cs
			if $info[10] = "Hewlett-Packard" Then
				_GUICtrlListView_SetItemText($list1, $i, _getHPPowerRecovery($compName), $iColBIOSPwr)
			EndIf
			#ce
		EndIf

		if $checkSw Then
			#cs
			$citrixInfo = _appUninstallFind("Citrix XenApp Plugin for Hosted Apps", $compName)
			If IsArray($citrixInfo) Then
				$citrix = $citrixInfo[1]
			Else
				$citrixInfo = _appUninstallFind("MetaFrame Presentation Server", $compName)
				If IsArray($citrixInfo) Then
					$citrix = $citrixInfo[1]
				Else
					$citrix = "Not found"
				EndIf
			EndIf
			dbout($compName & " Citrix clients searched", 1)
			if $citrix = "Not found" Then
				;{
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgCitrix, $iColSoftCitrix)
			ElseIf StringLeft($citrix, 4) = "11.0" Then
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColSoftCitrix)
			EndIf
			#ce

			$s92GUID = "{7A1FB67F-A340-472A-97C3-A6AFFE078AAE}"
			$s11GUID = "{388C130B-0079-46B4-A0D5-DC2DD7A89A7B}"
			$sVers = RegRead("\\" & $compName & "\HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & $s11GUID, "DisplayVersion")
			if NOT @error Then
				_GUICtrlListView_SetItem($hListView, $sVers, $i, $iColSoftCitrix, $lvImgCitrix)
			Else
				$sVers = RegRead("\\" & $compName & "\HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & $s92GUID, "DisplayVersion")
				if NOT @error Then
					_GUICtrlListView_SetItem($hListView, $sVers, $i, $iColSoftCitrix, $lvImgCitrix923)
				Else
					_GUICtrlListView_SetItem($hListView, "Nothing", $i, $iColSoftCitrix, $lvImgErr)
				EndIf
			EndIf
		EndIf

		if $checkAppV Then
			; name, version, last launch, pkg GUID, source OSD, global running count
			$aAppV = _AppvGetApps($compName)
			$err = @error


			if $err = 1 Then
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColInvision)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColAccessAnyware)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColStedmans)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColMicromedex)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColMIView)
			ElseIf $err = 2 Then
				_GUICtrlListView_SetItemText($hListView, $i, "No", $iColInvision)
				_GUICtrlListView_SetItemText($hListView, $i, "No", $iColAccessAnyware)
				_GUICtrlListView_SetItemText($hListView, $i, "No", $iColStedmans)
				_GUICtrlListView_SetItemText($hListView, $i, "No", $iColMicromedex)
				_GUICtrlListView_SetItemText($hListView, $i, "No", $iColMIView)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgWarn, $iColInvision)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgWarn, $iColAccessAnyware)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgWarn, $iColStedmans)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgWarn, $iColMicromedex)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgWarn, $iColMIView)
			Else
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColInvision)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColAccessAnyware)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColStedmans)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColMicromedex)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColMIView)
				For $iApp = 1 To $aAppV[0][0]
					if $aAppV[$iApp][2] = "1/1/1601 0:0" Then
						$aAppV[$iApp][2] = "[never]"
					EndIf

					if StringInStr($aAppV[$iApp][0], "Invision") Then
						_GUICtrlListView_SetItemText($hListView, $i, "Yes, launched on " & $aAppV[$iApp][2], $iColInvision)
						_GUICtrlListView_SetItemImage($hListView, $i, $lvImgOK, $iColInvision)
					EndIf
					If StringInStr($aAppV[$iApp][0], "AccessAnyware") Then
						_GUICtrlListView_SetItemText($hListView, $i, "Yes, launched on " & $aAppV[$iApp][2], $iColAccessAnyware)
						_GUICtrlListView_SetItemImage($hListView, $i, $lvImgOK, $iColAccessAnyware)
					EndIf
					If StringInStr($aAppV[$iApp][0], "Stedman") Then
						_GUICtrlListView_SetItemText($hListView, $i, "Yes, launched on " & $aAppV[$iApp][2], $iColStedmans)
						_GUICtrlListView_SetItemImage($hListView, $i, $lvImgOK, $iColStedmans)
					EndIf
					If StringInStr($aAppV[$iApp][0], "Micromedex") Then
						_GUICtrlListView_SetItemText($hListView, $i, "Yes, launched on " & $aAppV[$iApp][2], $iColMicromedex)
						_GUICtrlListView_SetItemImage($hListView, $i, $lvImgOK, $iColMicromedex)
					EndIf
					If StringInStr($aAppV[$iApp][0], "MI View") Then
						_GUICtrlListView_SetItemText($hListView, $i, "Yes, launched on " & $aAppV[$iApp][2], $iColMIView)
						_GUICtrlListView_SetItemImage($hListView, $i, $lvImgOK, $iColMIView)
					EndIf
				Next
			EndIf

			$user = RegRead("\\" & $compName & "\HKLM\Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName")
			if $user <> "" Then
				$userDN = _ADSamAccountNameToFQDN($user)
				if _ADIsMemberOf("CN=Clinical IDs THAM Wired,OU=Groups,OU=THR Users,DC=txhealth,DC=org", $userDN) Then
					_GUICtrlListView_SetItemText($hListView, $i, "Wired", $iColUserInAppV)
				ElseIf _ADIsMemberOf("CN=Clinical IDs THAM Wireless,OU=Groups,OU=THR Users,DC=txhealth,DC=org", $userDN) Then
					_GUICtrlListView_SetItemText($hListView, $i, "Wireless", $iColUserInAppV)
				Else
					_GUICtrlListView_SetItemText($hListView, $i, "No", $iColUserInAppV)
				EndIf
			EndIf

		EndIf

		If $checkCCM Then
			$iSCCM = _processExists("CcmExec.exe", $compName)
			if $iSCCM Then
				_GUICtrlListView_SetItemText($hListView, $i, "Yes", $iColCCM)
			Else
				_GUICtrlListView_SetItemText($hListView, $i, "No", $iColCCM)
			EndIf
		EndIf

		if $checkCert Then
			$aCert = _checkCert($compName)
			Local $err = @error, $ext = @extended
			if $err Then
				_GUICtrlListView_SetItemText($hListView, $i, "Missing", $iColCertSubject)
				_GUICtrlListView_SetItemImage($hListView, $i, $lvImgErr, $iColCertSubject)
				Switch $err
					Case 1 to 2
						dbout("Warning from " & $compName  & ": No valid certificate found.", 3)
					Case 4
						dbout("ERROR from " & $compName  & ": Registry keys didn't exist. SCCM may be broken.", 2)
					Case 3
						dbout("ERROR from " & $compName  & ": Couldn't connect to registry.", 2)
					Case 5
						dbout("ERROR from " & $compName & ": Certificate import error!", 2)
				EndSwitch
			Else
				_GUICtrlListView_SetItemText($hListView, $i, $aCert[0], $iColCertSubject)
				_GUICtrlListView_SetItemText($hListView, $i, $aCert[4], $iColCertValid)

				if StringInStr($aCert[0], $compName) Then
					_GUICtrlListView_SetItemImage($hListView, $i, $lvImgOK, $iColCertSubject)
				Else
					_GUICtrlListView_SetItemImage($hListView, $i, $lvImgWarn, $iColCertSubject)
				EndIf

				if $aCert[4] = True Then
					_GUICtrlListView_SetItemImage($hListView, $i, $lvImgOK, $iColCertValid)
				Else
					_GUICtrlListView_SetItemImage($hListView, $i, $lvImgWarn, $iColCertValid)
				EndIf
			EndIf
		EndIf

		if $checkMon Then
			$mInfo = _monitorInfoSemiSync($compName)
			If NOT @error Then
				_GUICtrlListView_SetItemText($hListView, $i, $mInfo[0], $iColMonitor)
				_GUICtrlListView_SetItemText($hListView, $i, $mInfo[1], $iColMSerial)
			Else
				_GUICtrlListView_SetItemText($hListView, $i, "EDID not found", $iColMSerial)
			EndIf
		EndIf

		_GUICtrlListView_SetItemChecked($hListView, $i, False)
		_GUICtrlListView_SetItemText($list1, $i, Round(TimerDiff($timer), 0), $iColTime)

		If $bStop Then
			$bStop = False
			ExitLoop
		EndIf
	Next

	For $iCol = 1 To _GUICtrlListView_GetColumnCount($list1)
		_GUICtrlListView_SetColumnWidth($list1, $iCol, $LVSCW_AUTOSIZE)
	Next
	upStatusBar($i-1 & " computers contacted in " & timeFormat(TimerDiff($timerTotal)))
	WinSetTitle($Form1, "", "remoteMassInfo " & $version)
EndFunc   ;==>tbQuery_Click

Func tbWMI_Click()
	$mark = _GUICtrlListView_GetSelectionMark($list1)
	If $mark <> -1 Then
		ShellExecute("remoteWmiInfo.exe", "/c " & _GUICtrlListView_GetItemText($list1, $mark, $iColName))
	Else
		upStatusBar("No computer selected!")
	EndIf
EndFunc   ;==>tbWMI_Click


#endregion ------------------------------ Toolbar items

Func upStatusBar($text)
	_GUICtrlStatusBar_SetText($statusBar2, _NowTime(5) & " >", 0)
	_GUICtrlStatusBar_SetText($statusBar2, $text, 2)
;~ 	GUICtrlSetData($progress, $iprogress)
EndFunc   ;==>upStatusBar

Func dbout($sMsg, $iLevel = 1, $sLine = @ScriptLineNumber)
	; info out by default, debuglevel 2 by default (errors, info)
	; debug w/ win, debug, warnings, errors, info, nothing

	if $debugLevel >= $iLevel Then
		If Not @Compiled Then
			ConsoleWrite(@HOUR & ":" & @MIN & ":" & @SEC & "." & numberPadZeroesInt(@MSEC, 3) & " (L" & $sLine & "): " & $sMsg & @CRLF)
		Else
			_GUICtrlEdit_AppendText($hConsoleEdit, @HOUR & ":" & @MIN & ":" & @SEC & "." & numberPadZeroesInt(@MSEC, 3) & ": " & $sMsg & @CRLF)
		EndIf
	EndIf
EndFunc   ;==>dbout

Func AutoItErr()
	$sErrText = "!COM Error: " & $oError.description & @CRLF
	$sErrText &= "windesc: " & $oError.windescription & @CRLF
	$sErrText &= "number: " & $oError.Number & @CRLF
	$sErrText &= "hex: " & Hex($oError.Number, 8) & @CRLF
	$sErrText &= "source: " & $oError.Source & @CRLF

	If Not @Compiled Then
		$sErrText &= "line #: " & $oError.ScriptLine & @CRLF
	EndIf

	dbout($sErrText, 2)

	SetError(69, $oError.Number)
EndFunc   ;==>AutoItErr

Func WM_GETMINMAXINFO($hWnd, $MsgID, $wParam, $lParam)
	#forceref $MsgID, $wParam
	Switch $hWnd
		Case WinGetHandle($Form1)
			Local $minmaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)

			DllStructSetData($minmaxinfo, 7, 500); min width
			DllStructSetData($minmaxinfo, 8, 200); min height
		Case WinGetHandle($Form2)
			Local $minmaxinfo = DllStructCreate("int;int;int;int;int;int;int;int;int;int", $lParam)

			DllStructSetData($minmaxinfo, 7, 250); min width
			DllStructSetData($minmaxinfo, 8, 200); min height
	EndSwitch
EndFunc   ;==>WM_GETMINMAXINFO

Func stringToBool($s)
	if $s = "True" Then
		Return True
	ElseIf $s = "False" Then
		Return False
	Else
		Return SetError(1)
	EndIf
EndFunc
