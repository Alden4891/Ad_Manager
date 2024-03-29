#AutoIt3Wrapper_AU3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
#include <AD.au3>

; Open Connection to the Active Directory
_AD_Open()
If @error Then Exit MsgBox(16, "Active Directory Example Skript", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)

Global $aProperties[1][2]

; *****************************************************************************
; Example 1
; Displays properties for the current user
; *****************************************************************************
$aProperties = _AD_GetObjectProperties(@UserName)
_ArrayDisplay($aProperties, "Active Directory Functions - Example 1 - Properties for user '" & @UserName & "'")

; *****************************************************************************
; Example 2
; Display only selected properties for the current user
; *****************************************************************************
$aProperties = _AD_GetObjectProperties(@UserName, "displayname,distinguishedName")
_ArrayDisplay($aProperties, "Active Directory Functions - Example 2 - Properties for user '" & @UserName & "'")

; *****************************************************************************
; Example 3
; Displays properties for the current computer
; *****************************************************************************
$aProperties = _AD_GetObjectProperties(@ComputerName & "$")
_ArrayDisplay($aProperties, "Active Directory Functions - Example 3 - Properties for computer '" & @ComputerName & "'")

; *****************************************************************************
; Example 4
; Get an array of group names the user is immediately a member of.
; Then display the properties of the first group.
; *****************************************************************************
Global $aUser = _AD_GetUserGroups(@UserName)
If $aUser[0] = 0 Then
	MsgBox(64, "Active Directory Functions - Example 4", "User '" & @UserName & "' is not a member of any group")
Else
	Global $sGroup = _AD_FQDNToSamAccountName($aUser[1])
	$aProperties = _AD_GetObjectProperties($sGroup)
	_ArrayDisplay($aProperties, "Active Directory Functions - Example 4 - Properties for group '" & $sGroup & "'")
EndIf

; *****************************************************************************
; Example 5
; Display the properties of the OU the group of example 4 is assigned to.
; *****************************************************************************
Global $sOU = StringTrimLeft($aUser[1], StringInStr($aUser[1], ","))
$aProperties = _AD_GetObjectProperties($sOU)
_ArrayDisplay($aProperties, "Active Directory Functions - Example 5 - Properties for OU '" & $sOU & "'")

; *****************************************************************************
; Example 6
; Queries a list of properties, some of them don't exist.
; _AD_GetObjectProperties ignores such properties returning a stripped down array.
; The example adds the non existing properties to the array and sets the value to "unset property".
; *****************************************************************************
Global $bFound, $sProperties = "aa,mail,xx,physicalDeliveryOfficeName,department,pwdLastSet,yy"
Global $aUserData = _AD_GetObjectProperties(@UserName, $sProperties)
_ArrayDisplay($aUserData, "Without unset properties")

$aProperties = StringSplit($sProperties, ",")
For $i = 1 To $aProperties[0]
	$bFound = False
	For $j = 1 To $aUserData[0][0]
		If $aUserData[$j][0] = $aProperties[$i] Then
			$bFound = True
			ExitLoop
		EndIf
	Next
	If $bFound = False Then
		_ArrayAdd($aUserData, $aProperties[$i] & "|unset property")
		$aUserData[0][0] = $aUserData[0][0] + 1
	EndIf
Next
_ArrayDisplay($aUserData, "With unset properties")

; Close Connection to the Active Directory
_AD_Close()
