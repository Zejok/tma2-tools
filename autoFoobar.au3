#AutoIt3Wrapper_icon=autoFoobar.ico

#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <buttonConstants.au3>

#include <_xmldomwrapper.au3>
#include <inet.au3>
#include <array.au3>

Global $PATH_SAVEDPOS = @MyDocumentsDir & "\Scripts\foobarPos.ini"

Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", $GUI_DOCKALL)
#Region ### START Koda GUI section ### Form=
$autoFoobar = GUICreate("autoFoobar", 213, 298)
GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_EVENT_CLOSE")

GUICtrlCreateLabel("Artist:", 8, 4, 37, 17)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")
$lArtist = GUICtrlCreateLabel("Not playing", 42, 4, 130, 17)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")

GUICtrlCreateLabel("Album:", 8, 16, 42, 17)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")
$lAlbum = GUICtrlCreateLabel("", 42, 16, 155, 17)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")

GUICtrlCreateLabel("Title:", 8, 28, 33, 17)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")
$lTitle = GUICtrlCreateLabel("", 42, 28, 155, 17)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")

GUICtrlCreateLabel("Misc:", 8, 40, 34, 17)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")
$lMisc = GUICtrlCreateLabel("", 42, 40, 155, 17)
GUICtrlSetFont(-1, 7, 400, 0, "Segoe UI")

$chkTop = GUICtrlCreateCheckbox("", 190, 5, 10, 15)
GUICtrlSetTip($chkTop, "Always on top?")

$Pic1 = GUICtrlCreatePic("", 8, 84, 196, 180, BitOR($SS_NOTIFY,$WS_GROUP,$WS_CLIPSIBLINGS))
$Progress1 = GUICtrlCreateProgress(8, 60, 164, 17)
$btnArt = GUICtrlCreateButton("", 178, 54, 24, 24, BitOR($BS_FLAT,$BS_ICON))
GUICtrlSetOnEvent(-1, "btnArt_Click")
GUICtrlSetTip(-1, "Show/hide album art.")
GUICtrlSetImage(-1, "C:\Users\TMA2\Documents\Scripts\Icons\Gnome-Go-Up.ico", 0, 0)

$btnSave = GUICtrlCreateButton("", 8, 276, 16, 16, BitOR($BS_ICON,$BS_ICON))
GUICtrlSetImage(-1, @MyDocumentsDir & "\Scripts\Icons\Gnome-Media-Floppy.ico", 0, 0)
GUICtrlSetOnEvent(-1, "btnSave_Click")
GUICtrlSetTip(-1, "Save playback position")
$btnRest = GUICtrlCreateButton("", 26, 276, 16, 16, BitOR($BS_ICON,$BS_ICON))
GUICtrlSetImage(-1, @MyDocumentsDir & "\Scripts\Icons\Gnome-Media-Skip-Forward.ico", 0, 0)
GUICtrlSetOnEvent(-1, "btnRest_Click")
GUICtrlSetTip(-1, "Restore playback position")

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $fb2k, $bShowArt = True, $iLength
setupTest()

$sTitleFormat = "%title% '('%bitrate%')'"
$sMiscFormat = "%codec% $if(%codec_profile%,'['%codec_profile%']',)'('%bitrate%'KBps)'"
; initial display
ConsoleWrite("Playing? " & $fb2k.Playback.Isplaying & @CRLF)

if $fb2k.Playback.IsPlaying Then
	playback_TrackChanged()
Else
	GUICtrlSetImage($Pic1, @ScriptDir & "\fb2k.jpg")
EndIf

While 1
	$bOnTop = GUICtrlRead($chkTop)

	Sleep(100)

	if $bOnTop <> GUICtrlRead($chkTop) Then
		if GUICtrlRead($chkTop) = $GUI_CHECKED Then
			WinSetOnTop($autoFoobar, "", 1)
		Else
			WinSetOnTop($autoFoobar, "", 0)
		EndIf
	EndIf
WEnd

Func GUI_EVENT_CLOSE()
	GUIDelete()
	FileDelete(@TempDir & "\autoFoobar-album.jpg")
	Exit
EndFunc

Func btnSave_Click()
	if $fb2k.Playback.IsPlaying Then
		IniWrite($PATH_SAVEDPOS, "Positions", $fb2k.Playback.FormatTitle("%filename_ext%"), $fb2k.Playback.Position)
	EndIf
EndFunc

Func btnRest_Click()
	if $fb2k.Playback.IsPlaying Then
		$fb2k.Playback.Seek(IniRead($PATH_SAVEDPOS, "Positions", $fb2k.Playback.FormatTitle("%filename_ext%"), $fb2k.Playback.Position))
	EndIf
EndFunc

Func btnArt_Click()
	Dim $temp = WinGetPos($autoFoobar)
	Dim $iSizeW = $temp[2], $iSizeH = $temp[3]
	Dim $iChange = 214
	Dim $iStep = 3
	Dim $pi = 3.14159265358979


	For $i=1 To $iChange Step $iStep
		$iTween = $i*($i/$iChange)
		Switch $bShowArt
			Case True
				WinMove($autoFoobar, "", $temp[0], $temp[1], $iSizeW, $iSizeH-$iTween)
			Case False
				WinMove($autoFoobar, "", $temp[0], $temp[1], $iSizeW, $iSizeH+$iTween)
		EndSwitch
;~ 		ConsoleWrite("motion:" & @TAB & $iTween & @CRLF)
	Next

	if $bShowArt = True Then
		GUICtrlSetImage($btnArt, "C:\Users\TMA2\Documents\Scripts\Icons\Gnome-Go-Down.ico", 0, 0)
	Else
		GUICtrlSetImage($btnArt, "C:\Users\TMA2\Documents\Scripts\Icons\Gnome-Go-Up.ico", 0, 0)
	EndIf

	$bShowArt = Not $bShowArt
EndFunc

Func setupTest()
	Dim $ProgID
	$ProgID = "Foobar2000.Application.0.7"
	ConsoleWrite("Looking for existing instance..." & @CRLF)
	$fb2k = ObjGet("", $ProgID)
	if Not IsObj($fb2k) then
		ConsoleWrite("Creating new instance..." & @CRLF)
		$fb2k = ObjCreate($ProgID)
		if Not IsObj($fb2k) then
			ConsoleWrite("Failed to get foobar2000 application object." & @CRLF)
			Exit
		endif
	endif
	ObjEvent($fb2k.Playback, "playback_")
EndFunc

Func playback_Started()
	ConsoleWrite("playback_Started()" & @CRLF)
EndFunc

Func playback_Stopped($iCode)
	Switch $iCode
		Case 0
			ConsoleWrite("playback_Stopped() : User stop" & @CRLF)
		Case 1
			ConsoleWrite("playback_Stopped() : EOF" & @CRLF)
		Case 2
			ConsoleWrite("playback_Stopped() : Switching tracks" & @CRLF)
		Case Else
			ConsoleWrite("playback_Stopped() : " & $iCode & @CRLF)
	EndSwitch
	GUICtrlSetImage($Pic1, @ProgramFilesDir & "\foobar2000\icons\generic.ico")
EndFunc

Func playback_Paused($i)
	if $i Then
		ConsoleWrite("playback_Paused()" & @CRLF)
	Else
		ConsoleWrite("playback_Paused() : Unpaused" & @CRLF)
	EndIf
EndFunc

Func playback_TrackChanged()
	With $fb2k.Playback
		$iLength = .Length
		$sArtist = .FormatTitle("%artist%")
		$sAlbum = .FormatTitle("%album%")
		$sTitle = .FormatTitle("%title%")
		$sMisc = .FormatTitle($sMiscFormat)
	EndWith

	GUICtrlSetData($lArtist, $sArtist)
	GUICtrlSetData($lAlbum, $sAlbum)
	GUICtrlSetData($lTitle, $sTitle)
	GUICtrlSetData($lMisc, $sMisc)

	$sPathArt = getAlbumArt()
	if $sPathArt Then
		GUICtrlSetImage($Pic1, $sPathArt)
	Else
		GUICtrlSetImage($Pic1, lfm_GetAlbumPic($sArtist, $sAlbum))
	EndIf
EndFunc

Func playback_PositionChanged($iPosition, $seeked)
	GUICtrlSetData($Progress1, Round(($iPosition/$iLength)*100))
EndFunc

Func getAlbumArt()
	$aTemp = StringSplit($fb2k.Playback.FormatTitle("%path%"), "\")
	$aTemp[$aTemp[0]] = "folder.jpg"

	$sDir = _ArrayToString($aTemp, "\", 1, $aTemp[0])
	ConsoleWrite($sDir & @CRLF)

	if FileExists($sDir) Then
		Return $sDir
	Else
		Return 0
	EndIf
EndFunc

Func lfm_GetAlbumPic($sArtist, $sAlbum)
	WinSetTitle($autoFoobar, "", "Fetching album art...")
	$sXML = _INetGetSource("http://ws.audioscrobbler.com/2.0/?method=album.getinfo&artist=" & $sArtist & "&album=" & $sAlbum & "&api_key=b25b959554ed76058ac220b7b2e0a026")
	_XMLLoadXML($sXML)
	$image = _XMLGetValue("/lfm/album/image[@size='extralarge']")
	$iSize = InetGetSize($image[1])
	$hGet = InetGet($image[1], @TempDir & "\autoFoobar-album.jpg", 0, 1)

	Do
		GUICtrlSetData($Progress1, Round((InetGetInfo($hGet, 0)/$iSize)*100))
;~ 		ConsoleWrite(Round((InetGetInfo($hGet, 0)/$iSize)*100) & ", ")
		Sleep(10)
	Until InetGetInfo($hGet, 2)

	InetClose($hGet)
	WinSetTitle($autoFoobar, "", "autoFoobar")
	Return @TempDir & "\autoFoobar-album.jpg"
EndFunc
