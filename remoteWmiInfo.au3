#AutoIt3Wrapper_Icon=Icons\TMA2\remoteWmiInfo.ico
#AutoIt3Wrapper_Res_Fileversion=0.7.11.2
#AutoIt3Wrapper_Res_Comment=Added checking for SCCM and fixed combobox history not being saved, as well as status icons for each computer. (you don't even wanna know how long it took)
#AutoIt3Wrapper_Res_Description=Remote WMI Info
#AutoIt3Wrapper_Res_FileVersion_AutoIncrement=Y
#AutoIt3Wrapper_Res_LegalCopyright=Jon Dunham
#AutoIt3Wrapper_Res_Field=Company|Mosaic Technology
#AutoIt3Wrapper_Res_Field=Legal Trademarks|Mosaic Technology

#AutoIt3Wrapper_Run_After=copy %out% "\\ftwgen01\this\field services\autoitscripts\remoteWmiInfo.exe"
#AutoIt3Wrapper_Change2CUI=N

Dim $compName, $go
Global $done = 0
Global $iDate = FileGetTime(@ScriptFullPath, 0, 1)
Global Const $version = FileGetVersion(@ScriptFullPath)

Global $debug = 0, $verbose = 1
Global $pathSave = @MyDocumentsDir & "\remoteWmiInfo Queries\"
Global $pathHistory = @MyDocumentsDir & "\rwiHistory.txt"
Global $pathSettings = @AppDataDir & "\TMA2 Tools\remoteWmiInfo.ini"

; FOR NEW PING METHOD
Global $oRefCount = ObjCreate("WbemScripting.SWbemNamedValueSet")
$oRefCount.Add("RefCount", 0)

Global Const $aMonAvail[18] = ["null", "Other", "Unknown", "Running or Full Power", "Warning", "In Test", "N/A", "Power Off", "Off Line", "Off Duty", "Degraded", "Not Installed", "Install Error", "Power Save - Unknown", "Power Save - Low Power Mode", "Power Save - Standby", "Power Cycle", "Power Save - Warning"]

; UDF
#include <date.au3>
#include <ie.au3>
#include <array.au3>
#include <misc.au3>
#include <guiStatusBar.au3>
#include <guiTab.au3>
#include <guiEdit.au3>
#include <guiToolbar.au3>
#include <guiImageList.au3>
#include <guiComboBoxEx.au3>

#include <guiListView.au3>

; Standard
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <Constants.au3>
#include <tabConstants.au3>
#include <treeViewConstants.au3>

; Personal
#include <_wirelessInfo.au3>
#include <_monitorInfo.au3>
#include <tma2.au3>
#include <_services.au3>
#include <userInfo.au3>

;~ #include "modernmenuraw.au3"

Opt("TrayAutoPause", 0)
Opt("GUICloseOnESC", 1)
OnAutoItExitRegister("onAutoItExit")
If $debug Then Opt("TrayIconDebug", 1)

;==============================================================================
;===== Check for Icons/Files
;==============================================================================


#region ### START Koda GUI section ### Form=
$frmInfo = GUICreate("Remote WMI Info " & $version, 337, 377, @DesktopWidth / 2, @DesktopHeight * 0.3)
$winSize = WinGetClientSize($frmInfo)


;==============================================================================
;===== Top Controls / Buttons
;==============================================================================

$cmbComp = _GUICtrlComboBoxEx_Create($frmInfo, "", 4, 30, 296, 150)

If FileRead($pathHistory) <> "" Then
	$aHistory = StringSplit(FileRead($pathHistory), "|")
	For $i = 1 To $aHistory[0]
		_GUICtrlComboBoxEx_AddString($cmbComp, $aHistory[$i])
	Next
EndIf
_GUICtrlComboBoxEx_SetEditTextMod($cmbComp, @ComputerName)

$hImage = _GUIImageList_Create(16, 16, 5, 1)
_GUIImageList_AddIcon($hImage, "Icons\Gnome\16\Network-Online.ico")
_GUIImageList_AddIcon($hImage, "Icons\Gnome\16\Network-Offline.ico")
_GUIImageList_AddIcon($hImage, "Icons\Gnome\16\Network-Error.ico")
_GUIImageList_AddIcon($hImage, "Icons\Gnome\Gnome-Colors-Emblem-Green.ico")
_GUIImageList_AddIcon($hImage, "Icons\Gnome\Gnome-Colors-Emblem-Desktop-Red.ico")
_GUICtrlComboBoxEx_SetImageList($cmbComp, $hImage)

$btnGo = GUICtrlCreateButton("&Query", 4, 4, 45, 22, BitOR($BS_DEFPUSHBUTTON, $BS_FLAT, $BS_ICON, $WS_GROUP))
GUICtrlSetImage($btnGo, "Icons\Gnome\16\Emblem-Default.ico", 0, 0)
GUICtrlSetTip($btnGo, "Query the specified computer." & @LF & "(Alt-Q)")
$btnSave = GUICtrlCreateButton("&Save", 54, 4, 45, 22, BitOR($BS_FLAT, $BS_ICON, $WS_GROUP))
GUICtrlSetImage($btnSave, "Icons\Gnome\16\Document-Save.ico", 0, 0)
GUICtrlSetTip($btnSave, "Save all fields to " & $pathSave & @LF & "(Alt-S)")
$btnRC = GUICtrlCreateButton("&Remote", 104, 4, 45, 22, BitOR($BS_FLAT, $BS_ICON, $WS_GROUP))
GUICtrlSetImage($btnRC, "Icons\Gnome\16\Preferences-Desktop-Remote-Desktop.ico", 0, 0)
GUICtrlSetTip($btnRC, "Start a remote control session." & @LF & "(Alt-R)")
$btnLO = GUICtrlCreateButton("&Logoff", 154, 4, 45, 22, BitOR($BS_FLAT, $BS_ICON, $WS_GROUP))
GUICtrlSetImage($btnLO, "Icons\Gnome\16\Application-Exit.ico", 0, 0)
GUICtrlSetTip($btnLO, "Log off the remote user." & @LF & "(Alt-L)")
$btnDef = GUICtrlCreateButton("&Default", 308, 30, 22, 20, BitOR($BS_FLAT, $BS_ICON, $WS_GROUP))
GUICtrlSetImage($btnDef, "Icons\Gnome\16\Edit-Undo.ico", 0, 0)
GUICtrlSetTip($btnDef, "Set target to the current computer." & @LF & "(Alt-D)")

$btnAbout = GUICtrlCreateButton("?", $winSize[0] - 18, 2, 16, 16, $BS_FLAT)
GUICtrlSetFont($btnAbout, 8.5, 1000)
;~ GUICtrlSetImage($btnAbout, "Icons\Gnome\16\Dialog-Question.ico", 0, 0)
GUICtrlSetTip($btnAbout, "About")


;==============================================================================
;===== GENERAL Tab
;==============================================================================

$Tab1 = GUICtrlCreateTab(4, 56, 330, 297, BitOR($TCS_RIGHTJUSTIFY, $TCS_MULTILINE))
GUICtrlSetResizing($Tab1, $GUI_DOCKWIDTH + $GUI_DOCKHEIGHT)
$TabSheet1 = GUICtrlCreateTabItem("General")

;==============================================================================
;===== COMPUTER Group
;==============================================================================
GUICtrlCreateGroup("Computer", 12, 78, 313, 138)

GUICtrlCreateLabel("Make / Model:", 20, 94, 80, 17)
$editMake = GUICtrlCreateInput("", 96, 94, 82, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
$editModel = GUICtrlCreateInput("", 182, 94, 136, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))

GUICtrlCreateLabel("Serial / Asset:", 20, 118, 80, 17)
$editSerial = GUICtrlCreateInput("", 96, 118, 112, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
$editAsset = GUICtrlCreateInput("", 212, 118, 106, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))

GUICtrlCreateLabel("BIOS / RAM:", 20, 142, 80, 17)
$editBIOS = GUICtrlCreateInput("", 96, 142, 112, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
$editRAM = GUICtrlCreateInput("", 212, 142, 106, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))

GUICtrlCreateLabel("CPU / Speed:", 20, 166, 80, 17)
$editCPU = GUICtrlCreateInput("", 96, 166, 162, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
$editCPUS = GUICtrlCreateInput("", 262, 166, 56, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))

GUICtrlCreateLabel("OS:", 20, 190, 80, 17)
$editOS = GUICtrlCreateInput("", 96, 190, 222, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))


;==============================================================================
;===== MONITOR Group
;==============================================================================

GUICtrlCreateGroup("Monitor", 12, 220, 313, 120)
GUICtrlCreateLabel("Model:", 20, 240, 70, 17)
GUICtrlCreateLabel("Serial:", 20, 264, 70, 17)
GUICtrlCreateLabel("Asset:", 20, 288, 70, 17)
GUICtrlCreateLabel("Description:", 20, 312, 70, 17)

$editDmName = GUICtrlCreateInput("", 96, 240, 196, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip(-1, "The model as reported by dumpEDID.", "Model")
$btnShowEDID = GUICtrlCreateButton("", 294, 240, 22, 22, BitOR($BS_ICON, $BS_FLAT))
GUICtrlSetImage($btnShowEDID, "Icons\Gnome\16\Video-Display2.ico", 0, 0)
GUICtrlSetTip($btnShowEDID, "Display full EDID information")
$editDmSerial = GUICtrlCreateInput("", 96, 264, 220, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip(-1, "The serial as reported by the EDID found in the registry." & @CRLF & "Only a portion of Dell serials are recorded there, so the prefix (e.g. CN0, MX0) is omitted." & @CRLF & "The vertical rule character ""|"" marks where the 5 middle characters would be normally.", "Monitor Serial")
$editDmAsset = GUICtrlCreateInput("", 96, 288, 196, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip(-1, "The asset tag recorded in Service Connect. Click the button to the right to try to get this information.", "Monitor Asset")
$btnGetDmAsset = GUICtrlCreateButton("", 294, 288, 22, 22, BitOR($BS_ICON, $BS_FLAT))
GUICtrlSetImage($btnGetDmAsset, "Icons\Gnome\Gnome-Video-Asset.ico", 0, 0)
$editDmDesc = GUICtrlCreateInput("", 96, 312, 220, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip(-1, "The PNP description that would be seen in Display Properties.", "Monitor title/description")
GUICtrlCreateGroup("", -99, -99, 1, 1)

;==============================================================================
;===== USERS Tab
;==============================================================================

$TabSheet4 = GUICtrlCreateTabItem("Users")

$hDefUserGroup = GUICtrlCreateGroup("Default User", 12, 88, 313, 126)
$editDUser = GUICtrlCreateEdit("", 15, 105, 306, 105, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY, $ES_WANTRETURN, $WS_VSCROLL))
GUICtrlCreateGroup("", -99, -99, 1, 1)

$hCurUserGroup = GUICtrlCreateGroup("Current User", 12, 216, 313, 126)
$editCUser = GUICtrlCreateEdit("", 15, 233, 306, 105, BitOR($ES_AUTOVSCROLL, $ES_AUTOHSCROLL, $ES_READONLY, $ES_WANTRETURN, $WS_VSCROLL))
GUICtrlCreateGroup("", -99, -99, 1, 1)

;==============================================================================
;===== NETWORK Tab
;==============================================================================

$TabSheet2 = GUICtrlCreateTabItem("Network")
GUICtrlCreateGroup("General", 12, 78, 313, 146)
GUICtrlCreateLabel("IP:", 20, 94, 70, 17)
$editIP = GUICtrlCreateInput("", 96, 94, 222, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlCreateLabel("MAC:", 20, 118, 70, 17)
$editMAC = GUICtrlCreateInput("", 96, 118, 222, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlCreateLabel("AMT GUID:", 20, 142, 60, 17)
$editGUID = GUICtrlCreateInput("", 96, 142, 222, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetFont($editGUID, 9, 400, 0, "Arial Narrow")
GUICtrlSetTip($editGUID, "The GUID (Globally Unique Identifier) provided by Intel AMT (Active Management Technology).")
GUICtrlCreateLabel("OU:", 20, 166, 70, 17)
$editDN = GUICtrlCreateInput("", 96, 166, 222, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlCreateLabel("Location:", 20, 190, 70, 17)
$editLoc = GUICtrlCreateInput("", 96, 190, 222, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip($editLoc, "The probable entity based on the first two IP octets.")
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Wireless", 12, 224, 313, 121)
GUICtrlCreateLabel("SSID:", 20, 243, 70, 17)
GUICtrlCreateLabel("BSSID:", 20, 267, 70, 17)
GUICtrlCreateLabel("Signal / Noise:", 20, 291, 70, 17)
GUICtrlCreateLabel("Channel:", 20, 315, 70, 17)

$editSSID = GUICtrlCreateInput("", 96, 243, 222, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip($editSSID, "The connected Service Set Identifier. In other words, the access point name.")
$editBSSID = GUICtrlCreateInput("", 96, 267, 222, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip($editBSSID, "The connected Base Service Set Identifier. In other words, the access point MAC address.")
$editSignal = GUICtrlCreateInput("", 96, 291, 106, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip($editSignal, "The wireless signal as reported by MSndis_80211_ReceivedSignalStrength.")
$editNoise = GUICtrlCreateInput("", 206, 291, 112, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip($editNoise, "The noise level from Atheros5000_NoiseFloor. This will only be available on computers with Atheros wireless chipsets.")
$editChannel = GUICtrlCreateInput("", 96, 315, 222, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))
GUICtrlSetTip($editChannel, "The wireless channel used.")
GUICtrlCreateGroup("", -99, -99, 1, 1)

;==============================================================================
;===== MISC Tab
;==============================================================================

$TabSheet3 = GUICtrlCreateTabItem("Software")

;==============================================================================
;===== SOFTWARE Group
;==============================================================================

$Group2 = GUICtrlCreateGroup("Software", 12, 80, 313, 264)

GUICtrlCreateLabel("IE:", 20, 100, 54, 17)
$chkIE = GUICtrlCreateCheckbox("", 80, 100, 17, 17)
GUICtrlSetTip(-1, "Check Internet Explorer version?")
GUICtrlSetState(-1, $GUI_CHECKED)
$editVerIE = GUICtrlCreateInput("", 104, 100, 212, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))

GUICtrlCreateLabel("McAfee:", 20, 124, 54, 17)
$chkMA = GUICtrlCreateCheckbox("", 80, 124, 17, 17)
GUICtrlSetTip(-1, "Check McAfee version & look for installed products?")
GUICtrlSetState(-1, $GUI_CHECKED)
$editVerMA = GUICtrlCreateEdit("", 104, 124, 212, 60, BitOR($ES_MULTILINE, $ES_AUTOHSCROLL, $ES_READONLY))

GUICtrlCreateLabel("LANDesk:", 20, 186, 54, 17)
$chkLD = GUICtrlCreateCheckbox("", 80, 186, 17, 17)
GUICtrlSetTip(-1, "Check LANDesk version?")
GUICtrlSetState(-1, $GUI_CHECKED)
$editLD = GUICtrlCreateInput("", 104, 186, 212, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))

GUICtrlCreateLabel("SCCM:", 20, 210, 54, 17)
$chkSCCM = GUICtrlCreateCheckbox("", 80, 210, 17, 17)
GUICtrlSetTip(-1, "Check for SCCM?")
GUICtrlSetState(-1, $GUI_CHECKED)
$editSCCM = GUICtrlCreateInput("", 104, 210, 212, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))

GUICtrlCreateLabel("Citrix:", 20, 234, 54, 17)
$chkCX = GUICtrlCreateCheckbox("", 80, 234, 17, 17)
GUICtrlSetTip(-1, "Check Citrix version?")
$editCitrix = GUICtrlCreateInput("", 104, 234, 212, 22, BitOR($ES_AUTOHSCROLL, $ES_READONLY))

$btnBrowse = GUICtrlCreateButton("", 104, 260, 25, 25, BitOR($BS_ICON, $BS_FLAT))
GUICtrlSetImage(-1, "Icons\Gnome\16\Folder-Remote.ico", 1, 0)
GUICtrlSetTip(-1, "Launch explorer pointed to the remote computer's C drive.", "Browse Computer", 1, 2)
GUICtrlSetState(-1, $GUI_DISABLE)

$btnSoftware = GUICtrlCreateButton("", 133, 260, 25, 25, BitOR($BS_ICON, $BS_FLAT))
GUICtrlSetImage(-1, "Icons\Gnome\16\System-Software-Installer.ico", 1, 0)
GUICtrlSetTip(-1, "Show all installed software.", "Software List", 1, 2)
GUICtrlSetState(-1, $GUI_DISABLE)


;==============================================================================
;===== AutoQuickbase
;==============================================================================

GUICtrlCreateTabItem("QB")
$btnQBFill = GUICtrlCreateButton("", 298, 318, 25, 25, BitOR($BS_FLAT, $BS_ICON, $WS_GROUP))
GUICtrlSetTip($btnQBFill, "Fill fields in open record edit window.")
GUICtrlSetImage($btnQBFill, "Icons\Gnome\16\Insert-Text.ico", 1, 0)
$btnQBDefault = GUICtrlCreateButton("", 270, 318, 25, 25, BitOR($BS_FLAT, $BS_ICON, $WS_GROUP))
GUICtrlSetTip($btnQBDefault, "Set default info & checkboxes.")
GUICtrlSetImage($btnQBDefault, "Icons\Gnome\16\View-Refresh.ico", 1, 0)

$chkQBTech = GUICtrlCreateCheckbox("Tech Name", 12, 85, 100, 17, BitOR($BS_CHECKBOX, $BS_AUTOCHECKBOX, $BS_RIGHTBUTTON, $WS_TABSTOP))
$chkQBTeam = GUICtrlCreateCheckbox("Team", 12, 108, 100, 17, BitOR($BS_CHECKBOX, $BS_AUTOCHECKBOX, $BS_RIGHTBUTTON, $WS_TABSTOP))
$inQBTech = GUICtrlCreateInput("", 120, 85, 200, 20)
$inQBTeam = GUICtrlCreateInput("", 120, 108, 200, 20)

#cs
	$chkQBInstalled =   GUICtrlCreateCheckbox("PC Installed",        12, 131, 81, 17)
	$chkQBTested =      GUICtrlCreateCheckbox("PC Tested",           12, 151, 81, 17)
	$chkQBNameCorrect = GUICtrlCreateCheckbox("PC Name Correct",     104, 131, 105, 17)
	$chkQBNIC =         GUICtrlCreateCheckbox("NIC set to 100/Half", 104, 151, 113, 17)
	$chkQBWireless =    GUICtrlCreateCheckbox("Wireless",            230, 131, 81, 17)
	$chkQBPrint =       GUICtrlCreateCheckbox("Print Queues",  230, 151, 113, 17)

	$chkQBSiteReady = 	GUICtrlCreateCheckbox("Site Ready", 12, 171,
	$chkQBSCUpdated
	$chkQBOvertime
	$chkQBHinder
#ce


$treeQB = GUICtrlCreateTreeView(12, 130, 160, 215, BitOR($TVS_HASBUTTONS, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS, $TVS_CHECKBOXES, $WS_GROUP, $WS_TABSTOP, $WS_BORDER))
$treeQB_0 = GUICtrlCreateTreeViewItem("Computer", $treeQB)
$chkQBMac = GUICtrlCreateTreeViewItem("MAC Address", $treeQB_0)
$chkQBSerial = GUICtrlCreateTreeViewItem("Service Tag", $treeQB_0)
$chkQBAsset = GUICtrlCreateTreeViewItem("Asset Tag", $treeQB_0)
$chkQBMSerial = GUICtrlCreateTreeViewItem("Monitor Serial", $treeQB_0)
$chkQBMAsset = GUICtrlCreateTreeViewItem("Monitor Asset", $treeQB_0)
$chkQBInstalled = GUICtrlCreateTreeViewItem("Installed", $treeQB_0)
$chkQBTested = GUICtrlCreateTreeViewItem("Tested", $treeQB_0)
$chkQBNameCorrect = GUICtrlCreateTreeViewItem("PC Name Correct", $treeQB_0)
$chkQBWireless = GUICtrlCreateTreeViewItem("Wireless", $treeQB_0)
$chkQBNIC = GUICtrlCreateTreeViewItem("NIC 100/Half", $treeQB_0)

$treeQB_10 = GUICtrlCreateTreeViewItem("Standard", $treeQB)
$chkQBSite = GUICtrlCreateTreeViewItem("Site Ready", $treeQB_10)
$chkQBServiceCenter = GUICtrlCreateTreeViewItem("Service Center", $treeQB_10)
$chkQBOvertime = GUICtrlCreateTreeViewItem("Overtime Install", $treeQB_10)
$chkQBHinder = GUICtrlCreateTreeViewItem("Hinder", $treeQB_10)

$chkQBPCDate = GUICtrlCreateCheckbox("PC Date", 180, 130, 100, 15)
$inQBPCDate = GUICtrlCreateInput(@MON & "-" & @MDAY & "-" & @YEAR, 180, 148, 140, 20)

$chkQBMDate = GUICtrlCreateCheckbox("Mount Date", 180, 170, 100, 15)
$inQBMDate = GUICtrlCreateInput(@MON & "-" & @MDAY & "-" & @YEAR, 180, 188, 140, 20)

$chkQB100Date = GUICtrlCreateCheckbox("100% Date", 180, 210, 100, 15)
$inQB100Date = GUICtrlCreateInput(@MON & "-" & @MDAY & "-" & @YEAR, 180, 228, 140, 20)

_defaultQB()
GUICtrlCreateTabItem("")

;==============================================================================
;===== LOG Tab
;==============================================================================

GUICtrlCreateTabItem("Log")
GUICtrlCreateGroup("Output", 12, 87, 313, 227)
$editLog = GUICtrlCreateEdit("", 16, 104, 303, 203, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $WS_VSCROLL))
GUICtrlSetBkColor(-1, 0x000000)
GUICtrlSetColor(-1, 0xAAFFFF)
GUICtrlSetFont(-1, 8, 400, 0, "Lucida Console")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$ctxLog = GUICtrlCreateContextMenu($editLog)
$ctxLogClear = GUICtrlCreateMenuItem("Clear", $ctxLog)
GUICtrlCreateMenuItem("", $ctxLog)
$ctxLogSave = GUICtrlCreateMenuItem("Save", $ctxLog)
$ctxLogNotepad = GUICtrlCreateMenuItem("Paste to Notepad", $ctxLog)
$chkVerbose = GUICtrlCreateCheckbox("Verbose Logging", 16, 318, 105, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$btnLogClear = GUICtrlCreateButton("", 298, 318, 25, 25, BitOR($BS_FLAT, $BS_ICON, $WS_GROUP))
GUICtrlSetTip($btnLogClear, "Clear log")
GUICtrlSetImage($btnLogClear, "Icons\Gnome\16\Edit-Clear.ico", 1, 0)
GUICtrlCreateTabItem("")

;==============================================================================
;===== SysMon Tab
;==============================================================================
#cs
GUICtrlCreateTabItem("Sm")
$oSysMon = ObjCreate("Sysmon.3" )
$hObjContainer = GUICtrlCreateObj($oSysMon, 12, 87, 313, 240)
GUICtrlCreateTabItem("")


$oSysMon.ShowValueBar = False
$oSysMon.ShowHorizontalGrid = True
$oSysMon.Counters.Add("\Process(*)\% Processor Time")
;~ $oSysMon.DisplayType=sysmonLineGraph
$oSysMon.GraphTitle="System Performance Overview"
#ce

;==============================================================================
;===== Status Bar
;==============================================================================

$statusBar = _GUICtrlStatusBar_Create($frmInfo)
; try to make a dynamically-sized status size
Dim $StatusBar1_PartsWidth[3] = [StringLen(_NowTime(5) & " >") * 6.5, 200, -1]
_GUICtrlStatusBar_SetParts($statusBar, $StatusBar1_PartsWidth)
_GUICtrlStatusBar_SetText($statusBar, _NowTime(5) & " >", 0)
_GUICtrlStatusBar_SetText($statusBar, "", 1)
_GUICtrlStatusBar_SetMinHeight($statusBar, 20)

#cs
$hMenu = GUICtrlCreateContextMenu()
_GUICtrlCreateODMenuItem("test1", $hMenu, "Icons\Gnome\16\Computer.ico")
_GUICtrlCreateODMenuItem("test2", $hMenu, "Icons\Gnome\16\Document-Open.ico")
_GUICtrlStatusBar_EmbedControl($statusBar, 2, $hMenu, 4)
#ce

#endregion ### END Koda GUI section ###

$oError = ObjEvent("AutoIt.Error", "TMA2_Error")
GUISetState(@SW_SHOW)

;==============================================================================
;===== Extra Info GUI
;==============================================================================

$frmExt = GUICreate("Extended Information", 450, 400, 100, 100, $WS_SIZEBOX, -1, $frmInfo)
$editExt = GUICtrlCreateEdit("", 0, 0, 450, 340, BitOR($WS_VSCROLL, $ES_WANTRETURN, $ES_READONLY))
GUICtrlSetResizing($editExt, $GUI_DOCKBORDERS)
GUICtrlSetFont($editExt, 8.5, 400, -1, "Lucida Console")
$btnExtClose = GUICtrlCreateButton("Close", 450 - 54, 400 - 56, 50, 24, $BS_DEFPUSHBUTTON)
GUICtrlSetResizing($btnExtClose, $GUI_DOCKSIZE + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM)

;==============================================================================
;===== Software GUI
;==============================================================================

$frmApps = GUICreate("Software List", 450, 400, 100, 100, $WS_SIZEBOX, -1, $frmInfo)
$lstApps = GUICtrlCreateListView("Name|Publisher|Version|Size|Date|Location|Uninstall String", 0, 0, 450, 340, -1, BitOR($LVS_EX_SNAPTOGRID,$LVS_EX_DOUBLEBUFFER,$LVS_EX_FULLROWSELECT))
GUICtrlSetResizing($lstApps, $GUI_DOCKBORDERS)
$btnLstClose = GUICtrlCreateButton("Close", 450 - 54, 400 - 56, 50, 24, $BS_DEFPUSHBUTTON)
GUICtrlSetResizing($btnLstClose, $GUI_DOCKSIZE + $GUI_DOCKRIGHT + $GUI_DOCKBOTTOM)

_GUICtrlListView_RegisterSortCallBack($lstApps)

;==============================================================================
;===== Arguments Check
;==============================================================================

If $cmdLine[0] Then
	Switch $cmdLine[1]
		Case "/c"
			_GUICtrlComboBoxEx_SetEditTextMod($cmbComp, $cmdLine[2])
			GUISetState(@SW_DISABLE)
			_go($cmdLine[2])
			GUISetState(@SW_ENABLE)
			WinActivate($frmInfo)
		Case Else
			MsgBox(4096, "remoteWmiInfo", "Argument " & $cmdLine[1] & " unknown.")
	EndSwitch
EndIf

;==============================================================================
;===== Main GUI Loop
;==============================================================================

Dim $goState, $logoffCheck

While 1
	$guiMsg = GUIGetMsg(1)

	Switch $guiMsg[0]
		Case $GUI_EVENT_CLOSE
			Switch $guiMsg[1]
				Case $frmInfo
					_GUICtrlListView_UnRegisterSortCallBack($lstApps)
					GUIDelete()
					Exit
				Case $frmExt
					GUISetState(@SW_HIDE, $frmExt)
				Case $frmApps
					GUISetState(@SW_HIDE, $frmApps)
			EndSwitch
		Case $btnDef
			_GUICtrlComboBoxEx_SetEditTextMod($cmbComp, @ComputerName)
		Case $btnSave
			_saveFields()
		Case $btnAbout
			aboutDiag()
		Case $btnBrowse
			$sComp = _GUICtrlComboBoxEx_GetEditTextMod($cmbComp)
			ShellExecute("explorer", "\\" & $sComp & "\C$")
		Case $btnGo
			GUISetState(@SW_DISABLE, $frmInfo)
			$sComp = _GUICtrlComboBoxEx_GetEditTextMod($cmbComp)
			if StringLeft($sComp, 5) = "call " Then
				if $debug AND @Compiled Then
					$aArgs = StringSplit(StringTrimLeft($sComp, 5), "|", 2)
					$aArgs[0] = "CallArgArray"

					$sFunc = $aArgs[1]
					_ArrayDelete($aArgs, 1)

					Call($sFunc, $aArgs)
				EndIf
			EndIf

			If $sComp <> "" Then
				$sComp = StringReplace($sComp, Chr(10), "")
				_GUICtrlComboBoxEx_SetEditTextMod($cmbComp, $sComp)
				_GUICtrlTab_SetCurFocus($Tab1, 5)
				$goState = _go($sComp)
				_GUICtrlTab_SetCurFocus($Tab1, 0)
			EndIf
			GUISetState(@SW_ENABLE, $frmInfo)
		Case $btnRC
			$sComp = _GUICtrlComboBoxEx_GetEditTextMod($cmbComp)
			If $sComp <> "" Then
				Select
					Case $goState = 1
						_LDstartRC($sComp)
					Case $goState = 2
						upStatus("Computer was not contactable.", 1)
					Case $goState = 0
						upStatus("Please query the computer first.", 1)
				EndSelect
			Else
				upStatus("Please enter a computer name.", 1)
			EndIf
		Case $btnShowEDID
			GUISetState(@SW_SHOW, $frmExt)
			WinSetTitle($frmExt, "", "Monitor EDID")
		Case $btnSoftware
			GUISetState(@SW_SHOW, $frmApps)
			_listSoftware(_GUICtrlComboBoxEx_GetEditTextMod($cmbComp))
		Case $btnLstClose
			GUISetState(@SW_HIDE, $frmApps)
		Case $lstApps
			_GUICtrlListView_SortItems($lstApps, GUICtrlGetState($lstApps))
		Case $btnGetDmAsset
			GUISetState(@SW_DISABLE, $frmInfo)
			$hTimerDm = TimerInit()
			$aDmInfo = _getSCInfo("*" & StringReplace(GUICtrlRead($editDmSerial), " | ", "*"), "asset.tag")
			If Not @error Then
				If $aDmInfo[0] <> "" Then
					GUICtrlSetData($editDmSerial, $aDmInfo[0])
					GUICtrlSetBkColor($editDmSerial, 0xBBEEBB)
				Else
					appendLog(_NowTime(5) & @TAB & "Monitor serial blank.")
				EndIf

				If StringLen($aDmInfo[0]) = 24 Then
					GUICtrlSetData($editDmSerial, StringReplace($aDmInfo[0], "-", ""))
					GUICtrlSetBkColor($editDmSerial, 0xBBEEBB)
				ElseIf StringLen($aDmInfo[0]) < 20 And StringLen($aDmInfo[0]) > 10 Then
					GUICtrlSetData($editDmSerial, $aDmInfo[0] & " [Bad length?]")
					GUICtrlSetBkColor($editDmSerial, 0xEEEE99)
				ElseIf StringLen($aDmInfo[0]) > 20 Then
					GUICtrlSetData($editDmSerial, $aDmInfo[0] & " [Bad length?]")
					GUICtrlSetBkColor($editDmSerial, 0xEEEE99)
				ElseIf $aDmInfo[0] = "" Then
					upStatus("Error getting monitor info; please retry.")
				Else
					GUICtrlSetData($editDmSerial, $aDmInfo[0])
					GUICtrlSetBkColor($editDmSerial, Default)
				EndIf

				If $aDmInfo[1] <> "" Then
					GUICtrlSetData($editDmAsset, $aDmInfo[1])
					GUICtrlSetBkColor($editDmAsset, 0xBBEEBB)
					upStatus("Monitor info retrieved successfully [" & Round(TimerDiff($hTimerDm)) & "ms]")
				Else
					GUICtrlSetBkColor($editDmAsset, 0xEEBBBB)
					upStatus("Couldn't get asset [" & Round(TimerDiff($hTimerDm)) & "ms]")
				EndIf
			ElseIf @error = 1 Then
				GUICtrlSetData($editDmAsset, "")
				upStatus("Couldn't get info (no records found).")
			ElseIf @error = 2 Then
				GUICtrlSetData($editDmAsset, "")
				upStatus("Couldn't get info (more than 1 result).")
			EndIf
			GUISetState(@SW_ENABLE, $frmInfo)
		Case $btnExtClose
			GUISetState(@SW_HIDE, $frmExt)
		Case $btnLO
			_logoffComp()
		Case $btnLogClear
			GUICtrlSetData($editLog, "")
		Case $btnQBDefault
			_defaultQB()
		Case $btnQBFill
			_fillQB()
	EndSwitch
	Sleep(10)
WEnd ; <== main loop

Func _go($compName)
;~ 	_filePush($pathHistory, $compName)
	$verbose = BitAND(GUICtrlRead($chkVerbose), $GUI_CHECKED)
	; Populate the array with WMI info, if possible
	$info = _wmiInfo($compName, $verbose)

	Local $iError = @error

	If $iError Then
		Switch $iError
			; Ping failed
			Case 7
				upStatus($info, 1)
				If $verbose Then appendLog(_NowTime(5) & @TAB & $info)
				;guiFlash($cmbComp, 0xFF0000, 100)

			; Computername contained illegal characters
			Case 8
				upStatus("Please enter a valid name.", 1)
				;guiFlash($cmbComp, 0xFF0000, 100)

			; Unable to get WMI info from computer after successful ping (this shouldn't really happen ever)
			Case 9
				upStatus("Unable to retrieve computer info after " & Round(@extended / 1000, 1) & "s.", 1)
				If $verbose Then appendLog(Round(@extended / 1000, 1) & "s:" & @TAB & "Unable to connect to WMI service.")
				;guiFlash($cmbComp, 0xFF0000, 100)
		EndSwitch
	EndIf

	if $iError Then
		Local $iImage = 1
		Local $iImage2 = 3
	Else
		Local $iImage = 0
		Local $iImage2 = 2
	EndIf


	$iPos = _GUICtrlComboBoxEx_FindStringExact($cmbComp, $compName)

	if $iPos = -1 Then
		_GUICtrlComboBoxEx_AddString($cmbComp, $compName, $iImage, $iImage2)
	Else
		_GUICtrlComboBoxEx_SetItemImage($cmbComp, $iPos, $iImage)
		_GUICtrlComboBoxEx_SetItemSelectedImage($cmbComp, $iPos, $iImage2)
	EndIf

	if $iError Then
		GUICtrlSetState($btnSoftware, $GUI_DISABLE)
		GUICtrlSetState($btnBrowse, $GUI_DISABLE)
		Return 2
	EndIf

	GUICtrlSetState($btnSoftware, $GUI_ENABLE)
	GUICtrlSetState($btnBrowse, $GUI_ENABLE)

	; Everything OK? DISPLAY THE INFO THEN DAMN
	$done = 1
	; ================================================= GENERAL
	$aTemp = StringSplit($info[0], ",", 2)
	GUICtrlSetData($editMake, $aTemp[0])
	GUICtrlSetData($editModel, $aTemp[1])
	GUICtrlSetData($editSerial, $info[1])
	GUICtrlSetData($editAsset, $info[10])
	GUICtrlSetData($editBIOS, $info[2])
	GUICtrlSetData($editCPU, StringReplace($info[33], "(R)", ""))
	GUICtrlSetData($editCPUS, $info[34] & " GHz")

	If $info[34] < 2 Then
		GUICtrlSetBkColor($editCPUS, 0xEEBBBB)
	Else
		GUICtrlSetBkColor($editCPUS, Default)
	EndIf

	GUICtrlSetData($editRAM, $info[3] & " GB")
	If $info[3] < 0.98 Then
		GUICtrlSetBkColor($editRAM, 0xEEBBBB)
	Else
		GUICtrlSetBkColor($editRAM, Default)
	EndIf

	GUICtrlSetData($editOS, StringReplace(StringReplace($info[4], "Microsoft", "MS"), "Service Pack", "SP"))
	If Not StringInStr($info[4], "XP") Then
		GUICtrlSetBkColor($editOS, 0xEEBBBB)
	Else
		GUICtrlSetBkColor($editOS, Default)
	EndIf
	; ========================================================
	; ================================================= MONITOR
	; ========================================================

	GUICtrlSetData($editDmName, $info[19])

	GUICtrlSetTip($editDmName, "The model as reported by dumpEDID." & @CRLF & @CRLF & _
			"Win32_DesktopMonitor.Availability = " & $info[23] & @CRLF & _
			"Description: " & $aMonAvail[$info[23]], "Model")

	GUICtrlSetData($editExt, $info[22])
	GUICtrlSetData($editDmSerial, $info[20])
	GUICtrlSetData($editDmAsset, "")
	GUICtrlSetData($editDmDesc, $info[21])

	; ================================================= USER
	; nt name, display name, description, job title, company, department, address, phone #, eMail
	$timer2 = TimerInit()
	; currentuser - $info[6]
	If $info[6] <> "" Then
		$uName = StringSplit($info[6], "\")
		$aUserInfo = _ADGetUserInfo("sAMAccountName", $uName[2])
		If IsArray($aUserInfo) Then
			If $verbose Then appendLog($info[5] + Round(TimerDiff($timer2), 0) & ":" & @TAB & "AD queried for " & $uName[2])

			GUICtrlSetData($editCUser, _
					"Full name:" & @TAB & $aUserInfo[1][1] & @CRLF & _
					"User name:" & @TAB & $uName[2] & @CRLF & _
					"Job title:" & @TAB & @TAB & $aUserInfo[1][3] & @CRLF & _
					"Description:" & @TAB & $aUserInfo[1][2] & @CRLF & _
					"Company:" & @TAB & @TAB & $aUserInfo[1][4] & @CRLF & _
					"Department:" & @TAB & $aUserInfo[1][5] & @CRLF & _
					"Address:" & @TAB & @TAB & $aUserInfo[1][6] & @CRLF & _
					"Phone:" & @TAB & @TAB & $aUserInfo[1][7] & @CRLF & _
					"Email:" & @TAB & @TAB & $aUserInfo[1][8] & @CRLF & _
					"Printer:" & @TAB & @TAB & $info[29])
			GUICtrlSetBkColor($editCUser, Default)
		Else
			appendLog($info[5] + Round(TimerDiff($timer2), 0) & ":" & @TAB & "Couldn't query AD.")
			GUICtrlSetData($editCUser, $info[6])
		EndIf
	Else
		GUICtrlSetData($editCUser, "No user currently logged on.")
		GUICtrlSetBkColor($editCUser, 0xEEBBBB)
	EndIf

	; defaultuser - $info[7]
	If $info[7] <> "" Then
		$duName = StringSplit($info[7], "\")
		$aDUserInfo = _ADGetUserInfo("sAMAccountName", $duName[2])
		If IsArray($aDUserInfo) Then
			If $verbose Then appendLog($info[5] + Round(TimerDiff($timer2), 0) & ":" & @TAB & "AD queried for " & $duName[2])

			GUICtrlSetData($editDUser, _
					"Full name:" & @TAB & $aDUserInfo[1][1] & @CRLF & _
					"User Name:" & @TAB & $duName[2] & @CRLF & _
					"Job title:" & @TAB & @TAB & $aDUserInfo[1][3] & @CRLF & _
					"Description:" & @TAB & $aDUserInfo[1][2] & @CRLF & _
					"Company:" & @TAB & @TAB & $aDUserInfo[1][4] & @CRLF & _
					"Department:" & @TAB & $aDUserInfo[1][5] & @CRLF & _
					"Address:" & @TAB & @TAB & $aDUserInfo[1][6] & @CRLF & _
					"Phone:" & @TAB & @TAB & $aDUserInfo[1][7] & @CRLF & _
					"Email:" & @TAB & @TAB & $aDUserInfo[1][8])
		Else
			GUICtrlSetData($editDUser, $info[7])
			appendLog($info[5] + Round(TimerDiff($timer2), 0) & ":" & @TAB & "Couldn't query AD.")
		EndIf
	EndIf

	; calc total time
	$info[5] += Round(TimerDiff($timer2), 0)
	upStatus($compName & " queried in " & $info[5] & "ms")

	; ================================================= NETWORK
	GUICtrlSetData($editIP, $info[11])
	GUICtrlSetData($editMAC, $info[9])
	GUICtrlSetData($editDN, $info[8])
	GUICtrlSetData($editLoc, $info[35])
	GUICtrlSetData($editGUID, $info[36])

	GUICtrlSetData($editSSID, $info[12])
	GUICtrlSetData($editBSSID, $info[18])
	GUICtrlSetData($editSignal, $info[13])
	GUICtrlSetData($editNoise, $info[14])
	GUICtrlSetData($editChannel, $info[15])

	; ================================================= Software

	GUICtrlSetData($editVerIE, $info[26])
	; Set field BkColor to red if below v7
	If StringLeft($info[26], 1) < 7 Then
		GUICtrlSetBkColor($editVerIE, 0xEECCCC)
	Else
		GUICtrlSetBkColor($editVerIE, Default)
	EndIf
	GUICtrlSetData($editVerMA, "Agent:" & @TAB & $info[31] & @CRLF & $info[32])
	; Set field BkColor to red if below v4
	If Number(StringLeft($info[31], 1)) < 4 Or $info[31] = "Not installed." Then
		GUICtrlSetBkColor($editVerMA, 0xEEEE99)
	Else
		GUICtrlSetBkColor($editVerMA, Default)
	EndIf

	GUICtrlSetData($editLD, $info[28])
	; Set field BkColor to red if not running.
	If $info[28] = "Not running." Then
		GUICtrlSetBkColor($editLD, 0xEECCCC)
	Else
		GUICtrlSetBkColor($editLD, Default)
	EndIf

	GUICtrlSetData($editCitrix, $info[30])
	; Set field BkColor to red if not found.
	If $info[30] = "Not found." Then
		GUICtrlSetBkColor($editCitrix, 0xEECCCC)
	Else
		GUICtrlSetBkColor($editCitrix, Default)
	EndIf

	If BitAND(GUICtrlRead($chkSCCM), $GUI_CHECKED) Then
		$iReturn = _processExists("CcmExec.exe", $compName)
		If $iReturn Then
			GUICtrlSetData($editSCCM, "Running (PID: " & $iReturn & ")")
			GUICtrlSetBkColor($editSCCM, Default)
		Else
			GUICtrlSetData($editSCCM, "Not running.")
			GUICtrlSetBkColor($editSCCM, 0xEECCCC)
		EndIf
	EndIf

	; flash window or edit box
	If Not WinActive($frmInfo) Then
		WinFlash($frmInfo, "", 4, 100)
	Else
		WinActivate($frmInfo)
		;guiFlash($cmbComp, 0x00FF00, 100)
	EndIf

	Return 1
EndFunc   ;==>_go

Func _wmiInfo($compName, $verbose = 0)
	$checkIE = BitAND(GUICtrlRead($chkIE), $GUI_CHECKED)
	$checkLD = BitAND(GUICtrlRead($chkLD), $GUI_CHECKED)
	$checkMA = BitAND(GUICtrlRead($chkMA), $GUI_CHECKED)
	$checkCX = BitAND(GUICtrlRead($chkCX), $GUI_CHECKED)

	; seterror if the computername string contains illegal characters

	If Not _computerNameLegal($compName) Then
		SetError(8)
		Return
	EndIf

	; init object variables
	Dim $oWMIService, $oAccount, $cItems
	Dim $dmEDID, $dmPNPDID, $dmName
	Dim $arrCU
	Dim $ping = Ping($compName, 1000)

	If @error Then
		Select
			Case @error = 1
				SetError(7)
				Return "Computer is offline."
			Case @error = 2
				SetError(7)
				Return "Computer is unreachable."
			Case @error = 3
				SetError(7)
				Return "Bad destination, please check the name."
			Case @error = 4
				SetError(7)
				Return "Problem contacting address."
		EndSelect
	EndIf

	appendLog("-[" & $compName & "]---------------")
	appendLog("Time:" & @TAB & _NowTime(5))
	appendLog("Ping:" & @TAB & $ping & "ms")



	; init arrays we'll return
	Dim $info[37], $mInfo, $infoW, $SID

	; get IP for no good damn reason
	TCPStartup()
	$info[11] = TCPNameToIP($compName)
	TCPShutdown()

	#cs
		167.99.0.0 = THFW (Fort Worth)
		172.22.0.0 = THAM Arlington Memorial
		--
		10.200.0.0 = Area 0
		10.201.0.0 = THFW (Fort Worth)
		10.202.0.0 = PHD (Denton)
		10.208.0.0 = THHEB (Hurst/Euless/Bedford)
		10.209.0.0 = HSW (Southwest Fort Worth)
		10.210.0.0 = THNW (Northwest/Azle)
		10.211.0.0 = THS (Stephenville)
		10.212.0.0 = THC (Cleburne)
		10.213.0.0 = Western
		10.214.0.0 = Fort Worth Remote Site
		10.215.0.0 = THAM (Arlington Memorial)
		10.224.0.0 = THP (Plano)
		10.225.0.0 = THA (Allen)
		10.226.0.0 = THK (Kaufman)
		10.227.0.0 = THW (Winnsboro)
		10.228.0.0 = PVN
		10.229.0.0 = Dallas Remote Site
		10.240.0.0 = THR HQ (Arlington)
	#ce

	$aIPSplit = StringSplit($info[11], ".")

	Switch Number($aIPSplit[1])
		Case 167
			$info[35] = "THFW (Fort Worth)"
		Case 172
			$info[35] = "THAM (Arlington Memorial)"
		Case 10
			Switch Number($aIPSplit[2])
				Case 200
					$info[35] = "Area 0"
				Case 201
					$info[35] = "THFW (Fort Worth)"
				Case 202
					$info[35] = "PHD (Dallas)"
				Case 208
					$info[35] = "THHEB (Hurst/Euless/Bedford)"
				Case 209
					$info[35] = "HSW (Southwest Fort Worth)"
				Case 210
					$info[35] = "THNW (Northwest/Azle)"
				Case 211
					$info[35] = "THS (Stephenville)"
				Case 212
					$info[35] = "THC (Cleburne)"
				Case 213
					$info[35] = "Western"
				Case 214
					$info[35] = "Fort Worth Remote Site"
				Case 215
					$info[35] = "THAM (Arlington Memorial)"
				Case 224
					$info[35] = "THP (Plano)"
				Case 225
					$info[35] = "THA (Allen)"
				Case 226
					$info[35] = "THK (Kaufman)"
				Case 227
					$info[35] = "THW (Winnsboro)"
				Case 228
					$info[35] = "PVN"
				Case 229
					$info[35] = "Dallas Remote Site"
				Case 240
					$info[35] = "THR HQ (Arlington)"
				Case Else
					$info[35] = "Unknown"
			EndSwitch
		Case Else
			$info[35] = "Unknown"
	EndSwitch

	; round off ping value
	$info[17] = Round($ping, -1)

	; start the response timer
	Dim $timer = TimerInit()

	; get the WMI object
	$oWMIService = ObjGet("winmgmts:\\" & $compName & "\root\cimv2")
	; check to see if the WMI object exists; if not, seterror and return
	; this check should now be deprecated due to the use of ping()
	If Not IsObj($oWMIService) Then
		SetError(9, TimerDiff($timer))
		Return
	EndIf

	appendLog("ms:" & @TAB & "Action")
	appendLog(" ")
	appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & ObjName($oWMIService) & " connected")
	; get defaultusername & software versions from the registry
	$info[7] = RegRead("\\" & $compName & "\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinLogon", "DefaultUserName")
	$sDDomain = RegRead("\\" & $compName & "\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinLogon", "DefaultDomainName")
;~ 	$info[36] = RegRead("\\" & $compName & "\HKLM\SOFTWARE\LANDesk\ManagementSuite\WinClient\AMT", "GUID")
	; IE
	If $checkIE Then $info[26] = RegRead("\\" & $compName & "\HKLM\SOFTWARE\Microsoft\Internet Explorer", "Version")
	; JRE
	$info[27] = RegRead("\\" & $compName & "\HKLM\SOFTWARE\JavaSoft\Java Runtime Environment", "CurrentVersion")
	; McAfee
	If $checkMA Then
		$info[31] = RegRead("\\" & $compName & "\HKLM\SOFTWARE\Network Associates\ePolicy Orchestrator\Application Plugins\EPOAGENT3000", "Version")
		If @error Then $info[31] = "Not installed."

		$i = 1
		While 1
			$regAvProd = RegEnumKey("\\" & $compName & "\HKLM\SOFTWARE\McAfee\ePolicy Orchestrator\Application Plugins", $i)
			If @error Then
				ExitLoop
			Else
				$sName = RegRead("\\" & $compName & "\HKLM\SOFTWARE\McAfee\ePolicy Orchestrator\Application Plugins\" & $regAvProd, "Product Name")
				$info[32] &= @CRLF & StringStripWS(StringReplace(StringReplace($sName, "Workstation", ""), "Module", ""), 7) & _
						" " & RegRead("\\" & $compName & "\HKLM\SOFTWARE\McAfee\ePolicy Orchestrator\Application Plugins\" & $regAvProd, "Version")
				$i += 1
			EndIf
		WEnd
		If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "McAfee products enumerated")
	EndIf

	; Citrix
	If $checkCX Then
		$citrixInfo = _appUninstallFind("Citrix XenApp Plugin for Hosted Apps", $compName)

		If IsArray($citrixInfo) Then
			$info[30] = $citrixInfo[1]
		Else
			$citrixInfo = _appUninstallFind("MetaFrame Presentation Server", $compName)
			If IsArray($citrixInfo) Then
				$info[30] = $citrixInfo[1]
			EndIf
		EndIf
		If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "Citrix clients searched")
	EndIf

	If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "Registry keys read")

	; LANDesk process (issuser)
	If $checkLD Then

		$info[28] = _processExists("issuser.exe", $compName)

		If Not $info[28] Then
			$info[28] = "Not running."
		Else
			$info[28] = "Running (PID: " & $info[28] & ")"
		EndIf

		If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "LANDesk process checked")
	EndIf

	;if $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "6 WQL queries executed")

	; get DN/OU
	$info[8] = _getDN($compName)
	If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "Computer OU retrieved")

	; get monitor info
	$mInfo = _monitorInfo($compName, 2)
	If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "Display device information processed")
	If IsArray($mInfo) Then
		$info[19] = $mInfo[0] ; model
		If $mInfo[1] <> "" Then
			$info[20] = $mInfo[1] ; serial
		Else
			$info[20] = $mInfo[4] ; num. serial if needed
		EndIf
		$info[21] = $mInfo[2] ; name
		$info[23] = $mInfo[5] ; availability
		If UBound($mInfo) = 7 Then _
				$info[22] = $mInfo[3] ; full EDID dump
	Else
		$info[19] = "No valid EDID found."
	EndIf

	; grab computer model & current user
	For $oItem In $oWMIService.InstancesOf("Win32_ComputerSystem")
		$info[0] = StringStripWS($oItem.Manufacturer, 2) & "," & $oItem.Model
		If $debug Then Debug_rWMI("Model: " & $info[0])
		$info[6] = $oItem.UserName
	Next

	; UUID
	For $oItem in $oWMIService.InstancesOf("Win32_ComputerSystemProduct")
		$info[36] = $oItem.UUID
	Next

	$info[36] = _stringConvertUUID($info[36])

	; get current user full name and description
	If $info[6] <> "" Then
		$arrCU = StringSplit($info[6], "\", 2)
		$oAccount = $oWMIService.Get('Win32_UserAccount.Name="' & $arrCU[1] & '",Domain="' & $arrCU[0] & '"')

		If IsObj($oAccount) Then
			$SID = $oAccount.SID
			If $debug Then Debug_rWMI($SID)
		EndIf
	EndIf

	; add domain to default username string
	$info[7] = $sDDomain & "\" & $info[7]

	; grab the serial and BIOS version
	For $oItem In $oWMIService.InstancesOf("Win32_BIOS")
		$info[1] = $oItem.SerialNumber
		$info[2] = $oItem.SMBIOSBIOSVersion
	Next

	; Asset tag
	For $oItem In $oWMIService.InstancesOf("Win32_SystemEnclosure")
		$info[10] = $oItem.SMBIOSAssetTag
	Next

	; HP Asset if needed
	If StringInStr($info[0], "HP Compaq 8000 Elite") And $info[10] = "" Or $info[10] = $info[1] Then
		If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "Asset blank or duped, trying HP CIM")
		Dim $oWMIService2 = ObjGet("winmgmts:{impersonationlevel=impersonate}//" & $compName & "/root/HP/InstrumentedBIOS")
		If Not IsObj($oWMIService2) Then
			$info[10] = "Blank or serial dupe"
		Else
			$temp = $oWMIService2.Get('HPBIOS_BIOSString.InstanceName="ACPI\\PNP0C14\\0_13"')
			$info[10] = StringStripWS($temp.Value, 8)

			If Not $info[10] Or StringIsSpace($info[10]) Then $info[10] = "Blank or serial dupe"
			$temp = 0
		EndIf

		$oWMIService2 = 0
	ElseIf $info[10] = "" Then
		$info[10] = "Blank or serial dupe"
	EndIf

	; Total RAM (GB)
	For $oItem In $oWMIService.InstancesOf("Win32_LogicalMemoryConfiguration")
		$info[3] = Round($oItem.TotalPhysicalMemory / 1048576, 2)
	Next

	; Processor
	$oCPU = $oWMIService.Get('Win32_Processor.DeviceID="CPU0"')
	If IsObj($oCPU) Then
		$info[33] = StringStripWS($oCPU.Name, 7)
		$info[34] = Round($oCPU.MaxClockSpeed / 1000, 2)
	EndIf

	; OS & Service Pack
	For $oItem In $oWMIService.InstancesOf("Win32_OperatingSystem")
		$info[4] = $oItem.Caption & " " & $oItem.CSDVersion
	Next

	; MAC for active NIC + check if IP is static or dynamic
	$cItems = $oWMIService.execquery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	For $oItem In $cItems
		$arrIP = $oItem.IPAddress
		If $arrIP[0] = $info[11] Then
			$info[9] = $oItem.MACAddress
			; check if static IP or not, indicating QS or otherwise special-purpose machine
			$info[16] = $oItem.DHCPEnabled
			If $info[16] = 0 Then
				$info[11] &= " (Static)"
			Else
				$info[11] &= " (Dynamic)"
			EndIf
		EndIf
	Next

	If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "" & ObjName($oWMIService) & " queries processed")

	; get wireless info if available
	$infoW = _wirelessInfo($compName)
	If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "Wireless statistics retrieved")

	; ssid
	$info[12] = $infoW[0]
	; signal
	$info[13] = $infoW[1]
	; noise
	$info[14] = $infoW[2]
	; channel
	$info[15] = $infoW[3]
	; BSSID
	$info[18] = $infoW[4]


	; Strip whitespace from end of model string, 'cause there's usually a lot
	$info[0] = StringStripWS($info[0], 2)

	#cs
	If $SID Then
		$info[29] = _defPrinterInfo($SID, $compName)
		If $verbose Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "Default printer retrieved")
	Else
	#ce

	$info[29] = _defPrinterInfo2($compName)
	If $verbose And Not @error Then appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & "Default printer retrieved")

	; get the TOTAL query time now
	$info[5] = Round(TimerDiff($timer), -1)

	Return $info
EndFunc   ;==>_wmiInfo

Func _saveFields()
	; Check that model field isn't blank
	If GUICtrlRead($editModel) = "" Then
		upStatus("No data to write!", 1)
		Return
	EndIf

	Dim $stateWrite, $file

	If Not FileExists($pathSave) Then
		DirCreate($pathSave)
	EndIf

	$stateWrite = FileWrite($pathSave & _GUICtrlComboBoxEx_GetEditTextMod($cmbComp) & "-" & @YEAR & @MON & @MDAY & ".txt", _
			"-( General )------------------------------" & @CRLF & _
			GUICtrlRead($editModel) & @CRLF & _
			GUICtrlRead($editSerial) & @CRLF & _
			GUICtrlRead($editBIOS) & @CRLF & _
			GUICtrlRead($editRAM) & @CRLF & _
			GUICtrlRead($editOS) & @CRLF & _
			GUICtrlRead($editMAC) & @CRLF & _
			GUICtrlRead($editIP) & @CRLF & _
			GUICtrlRead($editDN) & @CRLF & _
			GUICtrlRead($editCUser) & @CRLF & _
			GUICtrlRead($editDUser) & @CRLF & _
			"-( Wireless )-----------------------------" & @CRLF & _
			GUICtrlRead($editSSID) & @CRLF & _
			GUICtrlRead($editBSSID) & @CRLF & _
			GUICtrlRead($editSignal) & @CRLF & _
			GUICtrlRead($editNoise) & @CRLF & _
			GUICtrlRead($editChannel) & @CRLF & _
			"-( Monitor )------------------------------" & @CRLF & _
			GUICtrlRead($editDmName) & @CRLF & _
			GUICtrlRead($editDmSerial) & @CRLF & _
			GUICtrlRead($editDmDesc))

	If $stateWrite Then
		upStatus("Info saved to [ " & $pathSave & _GUICtrlComboBoxEx_GetEditTextMod($cmbComp) & "-" & @YEAR & @MON & @MDAY & ".txt ]")
	Else
		upStatus("Could not write file to [ " & $pathSave & _GUICtrlComboBoxEx_GetEditTextMod($cmbComp) & "-" & @YEAR & @MON & @MDAY & ".txt ]", 1)
	EndIf
EndFunc   ;==>_saveFields

Func _LDstartRC($sComp)
	$sRCPath = "C:\Program Files\LANDesk\ServerManager\RCViewer\isscntr.exe"
	If FileExists($sRCPath) Then
		;isscntr /a<address> /c<command> /l /s<core server>
		$sCore = "/sftwldmcor01"
		$sAddress = "/a" & $sComp
		$sCommand = '/c"remote control"'

		$sArgs = $sAddress & " " & $sCommand & " " & $sCore

		$iReturn = ShellExecute($sRCPath, $sArgs)
;~ 		appendLog("isscntr launched with return code: " & $iReturn)
		Return
	Else
		Dim $oIE = _IECreate("http://landesk/RemoteSession.aspx?machine=" & $sComp & "&operation=rc", 0, 0)
		_IEQuit($oIE)
		Return
	EndIf
EndFunc   ;==>_LDstartRC

Func _defaultQB()
	; check
	GUICtrlSetState($treeQB_0, $GUI_CHECKED)
	GUICtrlSetState($chkQBTech, $GUI_CHECKED)
	GUICtrlSetState($chkQBTeam, $GUI_CHECKED)
	GUICtrlSetState($chkQBMac, $GUI_CHECKED)
	GUICtrlSetState($chkQBSerial, $GUI_CHECKED)
	GUICtrlSetState($chkQBAsset, $GUI_CHECKED)
	GUICtrlSetState($chkQBMSerial, $GUI_CHECKED)
	GUICtrlSetState($chkQBMAsset, $GUI_CHECKED)
	GUICtrlSetState($chkQBInstalled, $GUI_CHECKED)
	GUICtrlSetState($chkQBTested, $GUI_CHECKED)
	GUICtrlSetState($chkQBNameCorrect, $GUI_CHECKED)
	; uncheck
	GUICtrlSetState($chkQBNIC, $GUI_UNCHECKED)
	GUICtrlSetState($chkQBWireless, $GUI_UNCHECKED)
	GUICtrlSetState($chkQBServiceCenter, $GUI_UNCHECKED)
	; check
	GUICtrlSetState($treeQB_10, $GUI_CHECKED)
	GUICtrlSetState($chkQBSite, $GUI_CHECKED)
	GUICtrlSetState($chkQBPCDate, $GUI_CHECKED)
	GUICtrlSetState($chkQB100Date, $GUI_CHECKED)
	; uncheck
	GUICtrlSetState($chkQBMDate, $GUI_UNCHECKED)
	GUICtrlSetState($chkQBOvertime, $GUI_UNCHECKED)
	GUICtrlSetState($chkQBHinder, $GUI_UNCHECKED)

	$aName = _ADGetUserInfo("sAMAccountName", @UserName)
	if IsArray($aName) Then
		$aName2 = StringSplit($aName[1][1], ", ", 3)
		GUICtrlSetData($inQBTech, $aName2[1] & " " & $aName2[0])
		GUICtrlSetData($inQBTeam, "Mosaic")
	EndIf


	GUICtrlSetData($inQBPCDate, @MON & "-" & @MDAY & "-" & @YEAR)
	GUICtrlSetData($inQBMDate, @MON & "-" & @MDAY & "-" & @YEAR)
	GUICtrlSetData($inQB100Date, @MON & "-" & @MDAY & "-" & @YEAR)

EndFunc   ;==>_defaultQB

Func _fillQB()
	If Not $done Then
		upStatus("Please query a computer first.")
		Return
	EndIf

	appendLog("")
	appendLog("Working...")
	Dim $hTimer = TimerInit()
	Dim $oShell = ObjCreate("Shell.Application")
	Dim $objIE

	For $objWindow In $oShell.Windows
;~ 		if StringInStr($objWindow.Document.Title, "AMH EPIC Deployments - Edit Device/Location Data") Then
		If StringInStr($objWindow.LocationURL, "quickbase.com") Then
			If StringInStr($objWindow.Document.Title, "AMH EPIC Deployments - Edit Device/Location Data") Then
				$objIE = $objWindow
				ExitLoop
			EndIf
		EndIf
	Next

	If Not IsObj($objIE) Then
		upStatus("Couldn't find Quickbase window.")
		Return SetError(1)
	EndIf

	$objForm = _IEFormGetObjByName($objIE, "editform")

;~ 		34 MAC
	If GUICtrlRead($chkQBMac) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_34")
		_IEFormElementSetValue($objElement, StringReplace(GUICtrlRead($editMAC), ":", ""))
	EndIf
;~ 		24 Service
	If GUICtrlRead($chkQBSerial) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_24")
		_IEFormElementSetValue($objElement, GUICtrlRead($editSerial))
	EndIf
;~ 		25  Asset Tag
	If GUICtrlRead($chkQBAsset) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_25")
		_IEFormElementSetValue($objElement, GUICtrlRead($editAsset))
	EndIf
;~ 		26  Monitor serial (possibly)
	If GUICtrlRead($chkQBMSerial) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_26")
		_IEFormElementSetValue($objElement, GUICtrlRead($editDmSerial))
		If StringInStr(GUICtrlRead($editDmSerial), "|") Then MsgBox(48, "Reminder", "Don't forget to fill in the FULL SERIAL!")
	EndIf
;~ 		27 Monitor asset (possibly)
	If GUICtrlRead($chkQBMAsset) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_27")
		_IEFormElementSetValue($objElement, GUICtrlRead($editDmAsset))
	EndIf
;~ 		70  Installed = 1
	If GUICtrlRead($chkQBInstalled) = $GUI_CHECKED Then
		_IEFormElementCheckBoxSelect($objForm, "0", "_fid_70", 1, "byIndex")
	Else
		_IEFormElementCheckBoxSelect($objForm, "0", "_fid_70", 0, "byIndex")
	EndIf
;~ 		76  PC Tested = 1
	If GUICtrlRead($chkQBTested) = $GUI_CHECKED Then
		_IEFormElementCheckBoxSelect($objForm, "0", "_fid_76", 1, "byIndex")
	Else
		_IEFormElementCheckBoxSelect($objForm, "0", "_fid_76", 0, "byIndex")
	EndIf
;~ 		73  NIC set = 'No'
	$objElement = _IEFormElementGetObjByName($objForm, "_fid_73")
	If GUICtrlRead($chkQBNIC) = $GUI_CHECKED Then
		_IEFormElementOptionSelect($objElement, "Yes")
	Else
		_IEFormElementOptionSelect($objElement, "No")
	EndIf
;~ 		57 Wireless = 'No'
	$objElement = _IEFormElementGetObjByName($objForm, "_fid_57")
	If GUICtrlRead($chkQBWireless) = $GUI_CHECKED Then
		_IEFormElementOptionSelect($objElement, "Yes")
	Else
		_IEFormElementOptionSelect($objElement, "No")
	EndIf
;~ 		78  Computer name correct = 1
	If GUICtrlRead($chkQBNameCorrect) = $GUI_CHECKED Then
		_IEFormElementCheckBoxSelect($objForm, "0", "_fid_78", 1, "byIndex")
	Else
		_IEFormElementCheckBoxSelect($objForm, "0", "_fid_78", 0, "byIndex")
	EndIf
;~ 		77  Install Tech Name
	If GUICtrlRead($chkQBTech) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_77")
		$aTemp = StringSplit(GUICtrlRead($inQBTech), " ")
		_IEFormElementOptionSelect($objElement, $aTemp[$aTemp[0]], 1, "bySearch")
	EndIf
;~ 		61  Site ready = 'Yes'
	$objElement = _IEFormElementGetObjByName($objForm, "_fid_61")
	If GUICtrlRead($chkQBSite) = $GUI_CHECKED Then
		_IEFormElementOptionSelect($objElement, "Yes")
	Else
		_IEFormElementOptionSelect($objElement, "No")
	EndIf
;~ 		75  Service Center = 'No'
	$objElement = _IEFormElementGetObjByName($objForm, "_fid_75")
	If GUICtrlRead($chkQBServiceCenter) = $GUI_CHECKED Then
		_IEFormElementOptionSelect($objElement, "Yes")
	Else
		_IEFormElementOptionSelect($objElement, "No")
	EndIf
;~ 		107 After hours = 'No'
	$objElement = _IEFormElementGetObjByName($objForm, "_fid_107")
	If GUICtrlRead($chkQBOvertime) = $GUI_CHECKED Then
		_IEFormElementOptionSelect($objElement, "Yes")
	Else
		_IEFormElementOptionSelect($objElement, "No")
	EndIf
;~ 		100 Hinder = 'No'
	$objElement = _IEFormElementGetObjByName($objForm, "_fid_100")
	If GUICtrlRead($chkQBHinder) = $GUI_CHECKED Then
		_IEFormElementOptionSelect($objElement, "Yes")
	Else
		_IEFormElementOptionSelect($objElement, "No")
	EndIf
;~ 		98  Team = 'Mosaic'
	If GUICtrlRead($chkQBTeam) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_98")
		_IEFormElementOptionSelect($objElement, GUICtrlRead($inQBTeam))
	EndIf
;~ 		63  PC Install Date
	If GUICtrlRead($chkQBPCDate) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_63")
		_IEFormElementSetValue($objElement, GUICtrlRead($inQBPCDate))
	EndIf
;~ 		64  Mount Install Date
	If GUICtrlRead($chkQBMDate) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_64")
		_IEFormElementSetValue($objElement, GUICtrlRead($inQBMDate))
	EndIf
;~ 		103 100% Complete Date
	If GUICtrlRead($chkQB100Date) = $GUI_CHECKED Then
		$objElement = _IEFormElementGetObjByName($objForm, "_fid_103")
		_IEFormElementSetValue($objElement, GUICtrlRead($inQB100Date))
	EndIf

	appendLog(@TAB & "done.")
	upStatus("Quickbase auto-completed in " & Round(TimerDiff($hTimer)) & "ms.")
EndFunc   ;==>_fillQB

Func _logoffComp()
	$sComp = _GUICtrlComboBoxEx_GetEditTextMod($cmbComp)
	If $sComp <> "" Then
		$sUser = _currentUser($sComp)
		If @error Or $sUser = "" Then Return

		$logoffCheck = MsgBox(51, "Remote WMI Info", "This will log off the current user: " & $sUser & "." & @CRLF & @CRLF & "Have you checked that the computer is not in use?")
		If $logoffCheck = 6 Then
			ShellExecute("\\ftwgen01\THIS\field services\autoItScripts\psTools\psshutdown.exe", "-o \\" & $sComp)
		EndIf
	Else
		MsgBox(0, "", "No computer entered or no user currently logged on.")
		Return
	EndIf
EndFunc   ;==>_logoffComp

Func _listSoftware($sComp)
	; DisplayName, Company, DisplayVersion, Size, InstallLocation, UninstallString
	Dim $i = 1, $iList = 0, $iTotalSize = 0
	Dim $hTimer = TimerInit()
	Dim $sRegLoc = "\\" & $sComp & "\HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall"
	if CPUArch($sComp) = "X64" Then
		Dim $sRegLoc64 = "\\" & $sComp & "\HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall"
	EndIf

	_GUICtrlListView_DeleteAllItems($lstApps)

	appendLog(Round($hTimer) & ":" & @TAB & "Enumerating software...")

	While 1
		$sEnum = RegEnumKey($sRegLoc, $i)
		if @error Then
			ExitLoop
		EndIf

		$i += 1
		$sKey = $sRegLoc & "\" & $sEnum

		$sName = RegRead($sKey, "DisplayName")
		if @error Then
			ContinueLoop
		ElseIf RegRead($sKey, "ParentKeyName") OR RegRead($sKey, "SystemComponent") Then
			ContinueLoop
		EndIf

;~ 		$iList += 1
		$sUninstall = RegRead($sKey, "UninstallString")
;~ 		if @error OR $sUninstall = "" Then ContinueLoop
		$sPub = RegRead($sKey, "Publisher")
		$sVer = RegRead($sKey, "DisplayVersion")
		$sLoc = RegRead($sKey, "InstallLocation")
		$sDate = RegRead($sKey, "InstallDate")
;~ 		$sIcon = RegRead($sKey, "DisplayIcon")
		$iSize = RegRead($sKey, "EstimatedSize")
		if $iSize Then
			$iTotalSize += $iSize
			$iSize = numberPadZeroesFloat(Round(Number($iSize)/1024, 2))
		Else
			$iSize = ""
		EndIf

;~ 		if @error Then ConsoleWrite("regread error: " & @error & @CRLF)
		$hItem = GUICtrlCreateListViewItem($sName & "|" & $sPub & "|" & $sVer & "|" & $iSize & "|" & $sDate & "|" & $sLoc & "|" & $sUninstall, $lstApps)
;~ 		$hItem = _GUICtrlListView_AddItem($ListView1, $sName, 0, _GUICtrlListView_GetItemCount($ListView1)+9999)
;~ 		_GUICtrlListView_SetItemText($lstApps, $iList, $sPub, 1)
;~ 		_GUICtrlListView_SetItemText($lstApps, $iList, $sVer, 2)
;~ 		_GUICtrlListView_SetItemText($lstApps, $iList, $iSize, 3)
;~ 		_GUICtrlListView_SetItemText($lstApps, $hItem, $sDate, 4)
;~ 		_GUICtrlListView_SetItemText($lstApps, $hItem, $sUninstall, 5)
;~ 		_GUICtrlListView_SetItemText($lstApps, $hItem, $sLoc, 4)


		#cs
		if $sIcon <> "" Then
			$aIcon = StringSplit($sIcon, ",")
			if $aIcon[0] > 1 Then
				$hIcon = _GUIImageList_AddIcon($hImage, $aIcon[1], $aIcon[2])
			Else
				$hIcon = _GUIImageList_AddIcon($hImage, $aIcon[1], 0)
			EndIf
			_GUICtrlListView_SetItemImage($ListView1, $hItem, $hIcon)
		Else
			_GUICtrlListView_SetItemImage($ListView1, $hItem, 0)
		EndIf
		#ce
	WEnd

	appendLog(Round(TimerDiff($hTimer)) & ":" & @TAB & "Enumerated " & $i & " items.")

	For $iCol=0 To _GUICtrlListView_GetColumnCount($lstApps)-2
		_GUICtrlListView_SetColumnWidth($lstApps, $iCol, $LVSCW_AUTOSIZE)
	Next

	_GUICtrlListView_SetColumn($lstApps, 0, "Name (" & _GUICtrlListView_GetItemCount($lstApps) & ")")
	_GUICtrlListView_SetColumn($lstApps, 3, "Size (" & numberPadZeroesFloat(Round(Number($iTotalSize)/1024, 2)) & ")", -1, 1)
EndFunc

Func upStatus($msg, $flash = 0, $color = 0x88DDCC)
	_GUICtrlStatusBar_SetText($statusBar, _NowTime(5) & " >", 0)
	_GUICtrlStatusBar_SetText($statusBar, $msg, 1)

	#CS
		if $flash Then
		;for $i=0 to 3
		;	_GUICtrlStatusBar_SetBkColor($statusBar, $color)
		;	Sleep(500)
		;	_GUICtrlStatusBar_SetBkColor($statusBar, $CLR_DEFAULT)
		;	Sleep(500)
		;Next
		Else
		_GUICtrlStatusBar_SetBkColor($statusBar, $CLR_MONEYGREEN)
		Sleep(250)
		_GUICtrlStatusBar_SetBkColor($statusBar, $CLR_DEFAULT)
		EndIf
	#ce
EndFunc   ;==>upStatus

Func appendLog($msg)
	; appendLog(Round(TimerDiff($timer), 0) & ":" & @TAB & msg)
	_GUICtrlEdit_AppendText($editLog, $msg & @CRLF)
EndFunc   ;==>appendLog

Func _filePush($sFilePath, $sName, $iLength = 10, $sDelim = "|")
	$hFile = FileOpen($sFilePath, 2)
	$sContents = FileRead($sFilePath)
	If $debug Then Debug_rWMI("_filePush $hFile: " & $hFile)

	$aContents = StringSplit($sContents, "|", 2)
	If $aContents[0] = $sName Then
		FileClose($hFile)
		Return 0
	EndIf

	If UBound($aContents) < $iLength Then
		If $sContents = "" Then
			$ret = FileWrite($hFile, $sName)
			If @error And $debug Then ConsoleWrite("_filePush Error: " & @error & @CRLF)
		Else
			$ret = FileWrite($hFile, $sName & "|" & $sContents)
			If @error And $debug Then ConsoleWrite("_filePush Error: " & @error & @CRLF)
		EndIf
	Else
		_ArrayPush($aContents, $sName, 1)
		$ret = FileWrite($hFile, _ArrayToString($aContents, $sDelim))
		If @error And $debug Then ConsoleWrite("_filePush Error: " & @error & @CRLF)
	EndIf
	If $debug Then ConsoleWrite("_filePush: " & $ret & ". Contents: " & FileRead($sFilePath) & @CRLF)
	FileClose($hFile)
EndFunc   ;==>_filePush

Func aboutDiag()
	GUISetState(@SW_DISABLE, $frmInfo)
	If @Compiled Then
		$ex = FileGetTime(@ScriptFullPath, 0, 1)
	Else
		$ex = "N/A"
	EndIf

	$qAbout = MsgBox(68, "About", "remoteWmiInfo " & $version & @CR & _
			"Compiled: " & StringLeft($ex, 8) & " - " & StringRight($ex, 6) & @CR & _
			"© Jon Dunham 2010" & @CR & _
			"dunham.jon@gmail.com" & @CR & @CR & _
			"Show readme file?")
	If $qAbout = 6 Then ShellExecute("wordpad", """\\ftwgen01\THIS\Field Services\autoItScripts\!readme.rtf""")
	GUISetState(@SW_ENABLE, $frmInfo)
	WinActivate($frmInfo)
EndFunc   ;==>aboutDiag

Func Debug_rWMI($sMsg, $sLine = @ScriptLineNumber)
	Local $sLineMsg = ""

	if @Compiled Then
		$sDate = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
		FileWriteLine(@AppDataDir & "\TMA2 Tools\" & @ScriptName & "-" & @ComputerName & "-Debug.log", $sDate & " ... " & $sMsg)
	Else
		ConsoleWrite(@ScriptName & " DEBUG (" & $sLine & ") ----- << " & @CRLF & $sMsg & @CRLF & "     >>" & @CRLF)
		$sLineMsg = "line " & $sLine
	EndIf

	if $verbose Then
		appendLog("DEBUG " & $sLineMsg & " : " & $sMsg)
	EndIf
EndFunc

Func TMA2_Error($sLine = @ScriptLineNumber)
	if @Compiled Then
		MsgBox(4144, "COM Error", "Number: " & $oError.Number & @CRLF & _
			"From: " & $oError.Source & @CRLF & _
			"Description: " & $oError.Description)
	Else
		Debug_rWMI("ERROR " & $oError.Number & " from " & $oError.Source & ": " & $oError.Description, $sLine)
	EndIf

EndFunc

Func onAutoItExit()
	$hCombo = _GUICtrlComboBoxEx_GetComboControl($cmbComp)
	if @error Then
		Debug_rWMI("Ex_GetComboControl (IshWnd " & IsHWnd($hCombo) & ") : " & $hCombo)
	EndIf

	$sCompHistory = _GUICtrlComboBox_GetList($hCombo)

	if $debug Then
		Local $sLB
		_GUICtrlComboBox_GetLBText($cmbComp, 0, $sLB)
		Debug_rWMI("ComboBox_GetCount : " & _GUICtrlComboBox_GetCount($cmbComp))
		Debug_rWMI("ComboBox_GetLBText : " & $sLB)
		Debug_rWMI("ComboBoxEx_GetList : " & $sCompHistory)
	EndIf


	$hFile = FileOpen($pathHistory, 2)
	If $hFile = -1 Then
		If @Compiled Then
			MsgBox(48, "Error", "FileOpen() error, cannot update list history." & @CRLF & "Attrib: " & FileGetAttrib($pathHistory))
		Else
			ConsoleWrite("FileOpen() error, attrib: " & FileGetAttrib($pathHistory) & @CRLF)
		EndIf
	Else
		FileWrite($hFile, $sCompHistory)
		FileClose($hFile)
	EndIf

	GUIDelete()
EndFunc   ;==>onAutoItExit