#include-once
#include <array.au3>

Func _wirelessInfo($sComputer = ".")

	dim $oWMIService = ObjGet( "winmgmts:\\" & $sComputer & "\root\WMI" )

	; Was the computer contactable?
	if NOT IsObj($oWMIService) Then
		Return 2
	EndIf

	dim $aSSID, $SSID, $BSSID, $signal, $noise, $channel, $info[5]
	Dim $raw
	; channel/frequency map
	; B/G =  1 -  14
	; A =   36 - 165
	Const $channels[27] =    ["1",      "2",      "3",      "4",      "5",      "6",      "7",      "8",      "9",      "10",     "11",     "12",     "13",     "14",     "36",     "40",     "44",     "48",     "52",     "56",     "68",     "60",     "149",    "153",    "157",    "161",    "165"]
	Const $frequencies[27] = ["2412000","2417000","2422000","2427000","2432000","2437000","2442000","2447000","2452000","2457000","2462000","2467000","2472000","2484000","5180000","5200000","5220000","5240000","5260000","5280000","5300000","5320000","5745000","5765000","5785000","5805000","5825000"]

	; SSID
	for $objWifi in $oWMIService.InstancesOf("MSNdis_80211_ServiceSetIdentifier")
		$aSSID = $objWifi.NDIS80211SSID
	Next
	
	; convert character array to a readable string
	For $i = 0 to UBound( $aSSID )-1
		if $aSSID[$i] < 32 OR $aSSID[$i] > 127 Then
			$aSSID[$i] = ""
		Else
			$aSSID[$i] = chr($aSSID[$i])
		EndIf
	Next
	$sSSID = _ArrayToString( $aSSID, "" )
	
	if NOT $sSSID Then
		$info[0] = "No wireless connection active."
		Return $info
	EndIf
	
	; Signal
	for $objWifi in $oWMIService.InstancesOf("MSNdis_80211_ReceivedSignalStrength")
		$signal = $objWifi.NDIS80211ReceivedSignalStrength & " dBm"
	Next

	; Noise
	for $objWifi in $oWMIService.InstancesOf("Atheros5000_NoiseFloor")
		$noise = -$objWifi.Value & " dBm"
	Next

	
	
	; AP MAC
	for $objWifi in $oWMIService.InstancesOf("MSNdis_80211_BaseServiceSetIdentifier")
		$BSSID = $objWifi.NDIS80211MacAddress
	Next
	
	; convert dec array to hex string
	For $i = 0 to UBound( $BSSID )-1
		$BSSID[$i] = Hex($BSSID[$i], 2)
	Next
	$BSSID = _ArrayToString( $BSSID, ":" )

	; Channel
	for $objWifi in $oWMIService.InstancesOf("MSNdis_80211_Configuration")
	   $raw = $objWifi.Ndis80211Config.DSConfig
	Next
   
	; find channel from frequency
	For $i=0 to ubound($frequencies)-1
		if $raw = $frequencies[$i] then
			$channel = $channels[$i]
		EndIf
	Next

	; Formatting (use stringreplace($info[4], ":", "") to remove or replace colons in the AP MAC if desired)

	

	$info[0] = $sSSID
	$info[1] = $signal
	$info[2] = $noise
	$info[3] = $channel; & " (" & $raw/1000000 & " GHz)"
	$info[4] = $BSSID

	Return $info

EndFunc
