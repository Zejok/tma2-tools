#include "tma2.au3"
#include-once

Global Const $EDID_HEADER = "0x00FFFFFFFFFFFF00"

Func _edidGetExt($EDID)
	; expects straight binary EDID

	; FF - serial	FE - string		FD - range limits
	; FC - model	FB - color		FA - timing data
	; F9 - undef.	0F - def. by manufacturer
	$aEDID = _byteArray($EDID)

	Local $return[4]

	For $offset = 54 To 108 Step 18
		if $aEDID[$offset] = 0x00 Then

			$block = BinaryMid($EDID, $offset+6, 12)
			$blockstr = BinaryToString($block)
			$iEnd = StringInStr($blockstr, @LF, 0, 1, 5)
			if $iEnd Then
				$string = StringLeft($blockstr, $iEnd-1)
			Else
				$string = $blockstr
			EndIf

			Switch $aEDID[$offset+3]
				Case 0xFF
					$return[1] = $string
				Case 0xFC
					$return[0] = $string
				Case 0xFE
					$return[2] = $string
			EndSwitch
		EndIf
	Next

	Return $return
EndFunc

Func _edidCheckValidity($EDID)
	; expects straight binary EDID
	if BinaryLen($EDID) <> 128 Then
		Return SetError(1)
	ElseIf String(BinaryMid($EDID, 1, 8)) <> $EDID_HEADER Then
		Return SetError(2)
	EndIf

	Local $checksum = 0x00

	$arrEDID = _byteArray($EDID)

	for $i = 0 to UBound($arrEDID)-1
		$checksum += $arrEDID[$i]
	Next

	Return NOT ($checksum)
EndFunc



Func __Test()
	; internal testing BS
	$hSearch = FileFindFirstFile("D:\Dox\code\EDID-related\test\edid.lcd*")

	While 1
		$search = FileFindNextFile($hSearch)
		if @error Then ExitLoop

		$hFile = FileOpen("D:\Dox\code\EDID-related\test\" & $search, 16)
		$EDID = FileRead($hFile)
		FileClose($hFile)

		$aEDID = _byteArray($EDID)
		ConsoleWrite("+ >> " & $search & @CRLF)
		ConsoleWrite("Valid? : " & _edidCheckValidity($EDID) & ", EDID Version: " & Int($aEDID[18]) & "." & Int($aEDID[19]) & @CRLF)

		$edidExt = _edidGetExt($EDID)
		For $i=0 To 3
			if $edidExt[$i] Then _
				ConsoleWrite("extended data: " & $edidExt[$i] & @CRLF)
		Next
	WEnd
EndFunc

#cs MISC REFERENCE CRAP
	...on PNPDeviceID...

	Q: Where do I obtain the ID Manufacturer Name?
	A: This information is obtained from Microsoft. Plug and Play device ID for a monitor/LCD consists of seven characters.
	   The first three characters represent the Vendor ID which is assigned by Microsoft. The four-character Product ID,
	   which is a 2-byte hexadecimal number, is assigned by the company producing the monitor. Companies can apply online
	   for a Vendor ID at http://www.microsoft.com/whdc/system/pnppwr/pnp/pnpid.mspx . Microsoft does not maintain Product IDs.
	   It is the individual company's responsibility to assure that they do not assign the same Product ID to two different devices.
		HIQ 6001 5 (&) 533723 f (&) 0 (&) UID257
		ACR 005E 5 (&) 533723 f (&) 0 (&) UID256
		ACR 005E 5 (&) 533723 f (&) 0 (&) UID260
		ACR 005E 5 (&) 533723 f (&) 0 (&) UID268 (&) 435460 ; INVALID?
		^   ^
		|   | PnP device ID
		|____ Manufacturer ID

		List of currently registered manufacturer PNPIDs
		http://download.microsoft.com/download/7/E/7/7E7662CF-CBEA-470B-A97E-CE7CE0D98DC2/ISA_PNPID_List.xlsx

======================( HEADER
	00–07 	Header information "00h FFh FFh FFh FFh FFh FFh 00h"
	08–09 	Manufacturer ID. These IDs are assigned by Microsoft. "00001=A”; “00010=B”; ... “11010=Z”.
				Bit 7 (at address 08h) is 0, the first character (letter) is located at bits 6 ? 2 (at address 08h),
				the second character (letter) is located at bits 1 & 0 (at address 08h) and bits 7 ? 5 (at address 09h),
				and the third character (letter) is located at bits 4 ? 0 (at address 09h).
	10–11 	Product ID Code (stored as LSB first). Assigned by manufacturer.
	12–15 	32-bit Serial Number. No requirement for the format. Usually stored as LSB first. In order to maintain compatibility
				with previous requirements the field should set at least one byte of the field to be non-zero if an ASCII
				serial number descriptor is provided in the detailed timing section.
	16 		Week of Manufacture. This varies by manufacturer. One way is to count January 1–7 as week 1, January 8–15 as
				week 2 and so on. Some count based on the week number (Sunday-Saturday). Valid range is 1-54.
	17 		Year of Manufacture. Add 1990 to the value for actual year.
	18 		EDID Version Number "01h"
	19 		EDID Revision Number "03h"


-	PnPID
		Type	Required or optional
		String	Required

		Plug and Play identifier string as found in the .inf file.
-	IdOriginal
		Type	Required or optional
		Boolean	Required

		Reserved for Microsoft. Must be True.
-	IsCompatibleID
		Type	Required or optional
		Boolean	Required

		Set to True if this is a compatible identifier as listed in the .inf file for this device.
-	ServiceName
		Type	Required or optional
		String	Required

		The device driver to be used for this device. Same as the string used under the HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services registry key.
-	UpperFilter
		Type	Required or optional
		Multistring	Optional

		Service names of upper filters that need to be installed for this device.
-	LowerFilter
		Type	Required or optional
		Multistring	Optional

  |-- struct-ish layout of sections
  |
  unsigned char sig[8];
  unsigned char mnft_id[2];            /* [8] manufaturer ID */
  unsigned char model_id[2];           /* [10] model ID */
  unsigned char ser_id[2];             /* [12] serial ID */
  unsigned char dummy_li[2];
  unsigned char week;                  /* [16] Week */
  unsigned char year;                  /* [17] + 1990 => Year */
  unsigned char major_version;         /* [18] */
  unsigned char minor_version;         /* [19] */
  unsigned char video_input_type;      /* [20] */
  unsigned char width;                 /* [21] */
  unsigned char height;                /* [22] */
  unsigned char gamma_factor;          /* [23] */
  unsigned char dpms;                  /* [24] */
  unsigned char rg;                    /* [25] colour information */
  unsigned char wb;                    /* [26] */
  unsigned char rY;                    /* [27] */
  unsigned char rX;                    /* [28] */
  unsigned char gY;                    /* [29] */
  unsigned char gX;                    /* [30] */
  unsigned char bY;                    /* [31] */
  unsigned char bX;                    /* [32] */
  unsigned char wY;                    /* [33] */
  unsigned char wX;                    /* [34] */
  unsigned char etiming1;              /* [35] */
  unsigned char etiming2;              /* [36] */
  unsigned char mtiming;               /* [37] */
  unsigned char stdtiming[16];         /* [38] */
  unsigned char text1[18];             /* [54/0x36] Product string */
  unsigned char text2[18];             /* [72/0x48] text 2 */
  unsigned char text3[18];             /* [90/0x5A] text 3 */
  unsigned char text4[18];             /* [108/0x6C] text 4 */
  unsigned char extension_blocks;      /* [126] number of following extensions*/
  unsigned char checksum;              /* [127] */
#ce

#cs REMOVED UNTIL I KNOW WHAT THE FUCK I'M DOING WITH DLLSTRUCTS AND POINTERS
; #STRUCTURE# ===================================================================================================================
; Name...........: $tagEDIDHDR
; Description ...: Defines the first 8-byte standard header
; Fields ........: [4]edidHeader			This should be checked to equal 0x00FFFFFFFFFFFF00

; Author ........: Jon Dunham (TMA-2)
; Remarks .......: Address 0x00 - 0x07
; ===============================================================================================================================
Global Const $tagEDIDHDR = "uint64 edidHeader;"

; #STRUCTURE# ===================================================================================================================
; Name...........: $tagEDIDPRODUCT
; Description ...: Defines the second 10-byte part describing manufacturer, model & assorted product information
; Fields ........: short manufacturerID		3 characters, assigned by MS. compressed 5bit ASCII codes.
;				   short productID			http://www.microsoft.com/whdc/system/pnppwr/pnp/pnpid.mspx / ISA PNPID
;				   int serialID
;				   byte manufactureWeek
;				   byte manufactureYear

;				   byte edidVersion
;				   byte edidRevision

; Author ........: Jon Dunham (TMA-2)
; Remarks .......:
; ===============================================================================================================================
Global Const $tagEDIDPRODUCT = $tagEDIDHDR & "short manufacturerID;" & _
											  "short productID;" & _
											  "int serialID;" & _
											  "byte manufactureWeek;" & _
											  "byte manufactureYear;" & _
											  "byte edidVersion;" &  _
											  "byte edidRevision"
#ce