#include-once
#include <array.au3>
#include <string.au3>

Func wirelessInfo($strComputer = ".")

	dim $objWMIService = ObjGet( "winmgmts:\\" & $strComputer & "\root\WMI" )

	; Was the computer contactable?
	if NOT IsObj($objWMIService) Then
		Return 2
	EndIf

	dim $SSID, $BSSID, $signal, $noise, $channel, $info[5]
	Dim $raw
	Const $channels[24] = ["1","2","3","4","5","6","7","8","9","10","11","40","36","44","48","52","56","60","64","149","153","157","161","165"]
	Const $frequencies[24] = ["2412000","2417000","2422000","2427000","2432000","2437000","2442000","2447000","2452000","2457000","2462000","5200000","5180000","5220000","5240000","5260000","5280000","5300000","5320000","5745000","5765000","5785000","5805000","5825000"]
	;$channels = _StringExplode($channels, ",")
	;$frequencies = _StringExplode($frequencies, ",")

	; Signal

	$colWifi = $objWMIService.ExecQuery( "Select * From MSNdis_80211_ReceivedSignalStrength where Active = True" )

	for $objWifi in $colWifi
		$signal = $objWifi.NDIS80211ReceivedSignalStrength & " dBm"
	Next

	; Noise

	$colWifi = $objWMIService.ExecQuery( "SELECT * FROM Atheros5000_NoiseFloor where Active = True")

	for $objWifi in $colWifi
		$noise = -$objWifi.Value & " dBm"
	Next

	; SSID

	$colWifi = $objWMIService.ExecQuery( "Select * From MSNdis_80211_ServiceSetIdentifier" )

	for $objWifi in $colWifi
		$SSID = $objWifi.NDIS80211SSID
	Next

	For $i = 0 to UBound( $SSID )-1
		if $SSID[$i] < 32 OR $SSID[$i] > 127 Then
			$SSID[$i] = ""
		Else
			$SSID[$i] = chr($SSID[$i])
		EndIf
	Next

	$SSID = _ArrayToString( $SSID, "" )

	; AP MAC

	$colWifi = $objWMIService.ExecQuery( "Select * From MSNdis_80211_BaseServiceSetIdentifier WHERE Active = True" )

	for $objWifi in $colWifi
		$BSSID = $objWifi.NDIS80211MacAddress
	Next

	For $i = 0 to UBound( $BSSID )-1
		$BSSID[$i] = Hex($BSSID[$i], 2)
	Next

	; Channel

	$colWifi = $objWMIService.ExecQuery("Select * From MSNdis_80211_Configuration WHERE Active = True" )

	for $objWifi in $colWifi
	   $raw = $objWifi.Ndis80211Config.DSConfig

	   for $i=0 to ubound($frequencies)-1
		  if $raw = $frequencies[$i] then
			 $channel = $channels[$i]
		  EndIf
		next

	next

	; Formatting (use stringreplace($info[4], ":", "") to remove or replace colons in the AP MAC if desired)

	$BSSID = _ArrayToString( $BSSID, ":" )

	$info[0] = $SSID
	$info[1] = $signal
	$info[2] = $noise
	$info[3] = $channel
	$info[4] = $BSSID

	Return $info

EndFunc
