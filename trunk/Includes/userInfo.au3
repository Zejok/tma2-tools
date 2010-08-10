#include-once

Func _ADGetUserInfo($type, $search)
	Dim $adoCommand, $adoConnection, $strBase, $strFilter, $strAttributes
	Dim $objRootDSE, $strDNSDomain, $strQuery, $adoRecordset

	; Setup ADO objects.
	$adoCommand = objCreate("ADODB.Command")
	$adoConnection = objCreate("ADODB.Connection")
	$adoConnection.Provider = "ADsDSOObject"
	$adoConnection.Open("Active Directory Provider")
	$adoCommand.ActiveConnection = $adoConnection

	; Search entire Active Directory domain.
	$objRootDSE = ObjGet("LDAP://RootDSE")

	if NOT IsObj($objRootDSE) Then Return SetError(2)

	$strDNSDomain = $objRootDSE.Get("defaultNamingContext")
	$strBase = "<LDAP://" & $strDNSDomain & ">"

	; Filter on user objects.
	$strFilter = "(&(objectCategory=person)(objectClass=user)(" & $type & "=" & $search & "))"

	; Comma delimited list of attribute values to retrieve.
	$strAttributes = "sAMAccountName,cn,description,title,company,department,streetAddress,telephoneNumber,mail"

	; Construct the LDAP syntax query.
	$strQuery = $strBase & ";" & $strFilter & ";" & $strAttributes & ";subtree"
	$adoCommand.CommandText = $strQuery
	$adoCommand.Properties("Page Size") = 100
	$adoCommand.Properties("Timeout") = 30
	$adoCommand.Properties("Cache Results") = False

	; Run the query.
	$adoRecordset = $adoCommand.Execute

	; nt name, display name, description, job title, company, department, address, phone #, eMail

	$count = $adoRecordset.RecordCount
	Dim $arrInfo[$count+1][9]
	$arrInfo[0][0] = $count
	$i=1

	$adoRecordset = $adoCommand.Execute

	; Enumerate the resulting recordset.
	if $count > 0 Then
		Do
			$arrInfo[$i][0] = $adoRecordset.Fields("sAMAccountName").Value
			$arrInfo[$i][1] = $adoRecordset.Fields("cn").value
			$arrInfo[$i][2] = $adoRecordset.Fields("description").value
			$arrInfo[$i][3] = $adoRecordset.Fields("title").value
			$arrInfo[$i][4] = $adoRecordset.Fields("company").value
			$arrInfo[$i][5] = $adoRecordset.Fields("department").value
			$arrInfo[$i][6] = $adoRecordset.Fields("streetAddress").value
			$arrInfo[$i][7] = $adoRecordset.Fields("telephoneNumber").value
			$arrInfo[$i][8] = $adoRecordset.Fields("mail").value

			; Move to the next record in the recordset.
			$adoRecordset.MoveNext
			$i += 1
		Until $adoRecordset.EOF
	Else
		SetError(1)
		Return 0
	EndIf

	; Clean up.
	$adoRecordset.Close
	$adoConnection.Close

	Return $arrInfo
EndFunc

Func _ADGetPrinterInfo($type, $search)
	Dim $adoCommand, $adoConnection, $strBase, $strFilter, $strAttributes
	Dim $objRootDSE, $strDNSDomain, $strQuery, $adoRecordset

	; Setup ADO objects.
	$adoCommand = objCreate("ADODB.Command")
	$adoConnection = objCreate("ADODB.Connection")
	$adoConnection.Provider = "ADsDSOObject"
	$adoConnection.Open("Active Directory Provider")
	$adoCommand.ActiveConnection = $adoConnection

	; Search entire Active Directory domain.
	$objRootDSE = ObjGet("LDAP://RootDSE")

	$strDNSDomain = $objRootDSE.Get("defaultNamingContext")
	$strBase = "<LDAP://" & $strDNSDomain & ">"

	; Filter on user objects.
	$strFilter = "(&(objectCategory=printQueue)(objectClass=printQueue)(" & $type & "=" & $search & "))"

	; Comma delimited list of attribute values to retrieve.
	$strAttributes = "name,location,shortServerName,uNCName"

	; Construct the LDAP syntax query.
	$strQuery = $strBase & ";" & $strFilter & ";" & $strAttributes & ";subtree"
	$adoCommand.CommandText = $strQuery
	$adoCommand.Properties("Page Size") = 100
	$adoCommand.Properties("Timeout") = 30
	$adoCommand.Properties("Cache Results") = False

	; Run the query.
	$adoRecordset = $adoCommand.Execute

	$count = $adoRecordset.RecordCount
	Dim $arrInfo[$count+1][4]
	$arrInfo[0][0] = $count
	$i=1

	$adoRecordset = $adoCommand.Execute

	; Enumerate the resulting recordset.
	if $count > 0 Then
		Do
			$arrInfo[$i][0] = $adoRecordset.Fields("name").Value
			$arrInfo[$i][1] = $adoRecordset.Fields("location").value
			$arrInfo[$i][2] = $adoRecordset.Fields("shortServerName").value
			$arrInfo[$i][3] = $adoRecordset.Fields("UNCName").value

			; Move to the next record in the recordset.
			$adoRecordset.MoveNext
			$i += 1
		Until $adoRecordset.EOF
	Else
		Return SetError(1)
	EndIf

	; Clean up.
	$adoRecordset.Close
	$adoConnection.Close

	Return $arrInfo
EndFunc