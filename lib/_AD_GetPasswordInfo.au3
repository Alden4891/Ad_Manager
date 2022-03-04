#include <AD.au3>

_AD_Open()
If @error Then Exit MsgBox(16, "Active Directory Example Skript", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)

; Get the password info
Global $aAD_PwdInfo[14][2] = [[13],["Maximum Password Age (days)"],["Minimum Password Age (days)"],["Enforce Password History (# of passwords remembered)"], _
		["Minimum Password Length"],["Account Lockout Duration (minutes)"],["Account Lockout Threshold (invalid logon attempts)"],["Reset account lockout counter after (minutes)"], _
		["Password last changed (YYYY/MM/DD HH:MM:SS local time)"],["Password expires (YYYY/MM/DD HH:MM:SS local time)"],["Password last changed (YYYY/MM/DD HH:MM:SS UTC)"], _
		["Password expires (YYYY/MM/DD HH:MM:SS UTC)"],["Password properties"],["Password expires - virtual property (YYYY/MM/DD HH:MM:SS local time)"]]

Global $aTemp = _AD_GetPasswordInfoEX()
For $iCount = 1 To $aTemp[0]
	$aAD_PwdInfo[$iCount][1] = $aTemp[$iCount]
Next
$aAD_PwdInfo[0][0] = $aTemp[0]

_ArrayDisplay($aAD_PwdInfo, "Active Directory Functions - Example 1")

_AD_Close()

; #FUNCTION# ====================================================================================================================
; Name...........: _AD_GetPasswordInfo
; Description ...: Returns password information retrieved from the domain policy and the specified user or computer account.
; Syntax.........: _AD_GetPasswordInfo([$sSamAccountName = @UserName])
; Parameters ....: $sObject - [optional] User or computer account to get password info for (default = @UserName). Format is sAMAccountName or FQDN
; Return values .: Success - Returns a one-based array with the following information:
;                  |1 - Maximum Password Age (days)
;                  |2 - Minimum Password Age (days)
;                  |3 - Enforce Password History (# of passwords remembered)
;                  |4 - Minimum Password Length
;                  |5 - Account Lockout Duration (minutes). 0 means the account has to be unlocked manually by an administrator
;                  |6 - Account Lockout Threshold (invalid logon attempts)
;                  |7 - Reset account lockout counter after (minutes)
;                  |8 - Password last changed (YYYY/MM/DD HH:MM:SS in local time of the calling user) or "1601/01/01 00:00:00" (means "Password has never been set")
;                  |9 - Password expires (YYYY/MM/DD HH:MM:SS in local time of the calling user) or empty when password has not been set before or never expires
;                  |10 - Password last changed (YYYY/MM/DD HH:MM:SS in UTC) or "1601/01/01 00:00:00" (means "Password has never been set")
;                  |11 - Password expires (YYYY/MM/DD HH:MM:SS in UTC) or empty when password has not been set before or never expires
;                  |12 - Password properties. Part of Domain Policy. A bit field to indicate complexity / storage restrictions
;                  |      1 - DOMAIN_PASSWORD_COMPLEX
;                  |      2 - DOMAIN_PASSWORD_NO_ANON_CHANGE
;                  |      4 - DOMAIN_PASSWORD_NO_CLEAR_CHANGE
;                  |      8 - DOMAIN_LOCKOUT_ADMINS
;                  |     16 - DOMAIN_PASSWORD_STORE_CLEARTEXT
;                  |     32 - DOMAIN_REFUSE_PASSWORD_CHANGE
;                  |13 - Calculated password expiration date/time. Identical with element 9 of this array.
;                  |     Returns a value even when fine grained password policy is in use; which means that most of the other elements of this array are blank or 0.
;                  |     This is a Virtual Attribute (aka "Pseudo Attribute", "Constructed Attribute" or "Back-link") where the value is calculated by the LDAP Server Implementation and is not actually part of the LDAP Entry.
;                  Failure - "", sets @error to:
;                  |1 - $sObject not found
;                  |2 - Error accessing the userAccountControl property. Might be caused by missing permissions. @extended is set to the COM error code
;                  Warning - Returns a one-based array (see Success), sets @extended to one of the following values (can be a combination of the following values e.g. 3 = 1 (Password does not expire) + 2 (Password has never been set)
;                  |1 - Password does not expire (User Access Control - UAC - is set)
;                  |2 - Password has never been set
;                  |4 - The Maximum Password Age is set to 0 in the domain. Therefore, the password does not expire
;                  |8 - The version of the accessed DC (needs to be >= 2008) does not support property MSDS-UserPasswordExpiryTimeComputed. Element 13 of the returned array is set to element 9.
;                  |16 - Function _AD_GetObjectProperties returned an error when querying property MSDS-UserPasswordExpiryTimeComputed. The error is ignored and element 13 of the returned array is set to element 9.
; Author ........: water
; Modified.......:
; Remarks .......: For details about password properties please check: http://msdn.microsoft.com/en-us/library/aa375371(v=vs.85).aspx
; Related .......: _AD_IsPasswordExpired, _AD_GetPasswordExpired, _AD_GetPasswordDontExpire, _AD_SetPassword, _AD_DisablePasswordExpire, _AD_EnablePasswordExpire, _AD_EnablePasswordChange,  _AD_DisablePasswordChange
; Link ..........: http://www.autoitscript.com/forum/index.php?showtopic=86247&view=findpost&p=619073, https://www.itprotoday.com/windows-78/jsi-tip-8294-how-can-i-return-domain-password-policy-attributes
; Example .......: Yes
; ===============================================================================================================================
Func _AD_GetPasswordInfoEX($sObject = @UserName)

	If $sObject = Default Then $sObject = @UserName
	If _AD_ObjectExists($sObject) = 0 Then Return SetError(1, 0, "")
	If StringMid($sObject, 3, 1) <> "=" Then $sObject = _AD_SamAccountNameToFQDN($sObject) ; sAMAccountName provided
	Local $iExtended = 0, $aPwdInfo[14] = [13], $oObject, $oUser, $sPwdLastChanged, $iUAC, $aTemp
	$oObject = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sAD_DNSDomain)
	ConsoleWrite("maxPwdAge: ")
	$aPwdInfo[1] = Int(__AD_Int8ToSecEX($oObject.Get("maxPwdAge"))) / 86400 ; Convert to Days
	If Int($aPwdInfo[1]) = -10675199 Then $aPwdInfo[1] = 0
	ConsoleWrite("minPwdAge: ")
	$aPwdInfo[2] = __AD_Int8ToSecEX($oObject.Get("minPwdAge")) / 86400 ; Convert to Days
	$aPwdInfo[3] = $oObject.Get("pwdHistoryLength")
	$aPwdInfo[4] = $oObject.Get("minPwdLength")
	; Account lockout duration: http://msdn.microsoft.com/en-us/library/ms813429.aspx
	; http://www.autoitscript.com/forum/topic/158419-active-directory-udf-help-support-iii/page-5#entry1173322
	ConsoleWrite("lockoutDuration: ")
	$aPwdInfo[5] = __AD_Int8ToSecEX($oObject.Get("lockoutDuration")) / 60 ; Convert to Minutes
	If $aPwdInfo[5] < 0 Or $aPwdInfo[5] > 99999 Then $aPwdInfo[5] = 0
	$aPwdInfo[6] = $oObject.Get("lockoutThreshold")
	ConsoleWrite("lockoutObservationWindow: ")
	$aPwdInfo[7] = __AD_Int8ToSecEX($oObject.Get("lockoutObservationWindow")) / 60 ; Convert to Minutes
	$oUser = __AD_ObjGet("LDAP://" & $sAD_HostServer & "/" & $sObject)
	$sPwdLastChanged = $oUser.Get("PwdLastSet")
	$iUAC = $oUser.userAccountControl
	If @error Then Return SetError(2, @error, "")
	; Has user account password been changed before?
	If $sPwdLastChanged.LowPart = 0 And $sPwdLastChanged.HighPart = 0 Then
		$iExtended = BitOR($iExtended, 2)
		$aPwdInfo[8] = "1601/01/01 00:00:00"
		$aPwdInfo[10] = "1601/01/01 00:00:00"
	Else
		Local $sTemp = DllStructCreate("dword low;dword high")
		DllStructSetData($sTemp, "Low", $sPwdLastChanged.LowPart)
		DllStructSetData($sTemp, "High", $sPwdLastChanged.HighPart)
		; Have to convert to SystemTime because _Date_Time_FileTimeToStr has a bug (#1638)
		Local $sTemp2 = _Date_Time_FileTimeToSystemTime(DllStructGetPtr($sTemp))
		$aPwdInfo[10] = _Date_Time_SystemTimeToDateTimeStr($sTemp2, 1)
		; Convert PwdlastSet from UTC to Local Time
		$sTemp2 = _Date_Time_SystemTimeToTzSpecificLocalTime(DllStructGetPtr($sTemp2))
		$aPwdInfo[8] = _Date_Time_SystemTimeToDateTimeStr($sTemp2, 1)
		; Is user account password set to expire?
		If BitAND($iUAC, $ADS_UF_DONT_EXPIRE_PASSWD) = $ADS_UF_DONT_EXPIRE_PASSWD Or $aPwdInfo[1] = 0 Then
			If BitAND($iUAC, $ADS_UF_DONT_EXPIRE_PASSWD) = $ADS_UF_DONT_EXPIRE_PASSWD Then $iExtended = BitOR($iExtended, 1)
			If $aPwdInfo[1] = 0 Then $iExtended = BitOR($iExtended, 4) ; The Maximum Password Age is set to 0 in the domain. Therefore, the password does not expire
		Else
			$aPwdInfo[11] = _DateAdd("d", $aPwdInfo[1], $aPwdInfo[10])
			$sTemp2 = _Date_Time_EncodeSystemTime(StringMid($aPwdInfo[11], 6, 2), StringMid($aPwdInfo[11], 9, 2), StringMid($aPwdInfo[11], 1, 4), StringMid($aPwdInfo[11], 12, 2), StringMid($aPwdInfo[11], 15, 2), StringMid($aPwdInfo[11], 18, 2))
			; Convert PasswordExpires from UTC to Local Time
			$sTemp2 = _Date_Time_SystemTimeToTzSpecificLocalTime(DllStructGetPtr($sTemp2))
			$aPwdInfo[9] = _Date_Time_SystemTimeToDateTimeStr($sTemp2, 1)
		EndIf
	EndIf
	$aPwdInfo[12] = $oObject.Get("pwdProperties")
	$aTemp = _AD_GetObjectProperties($sObject, "MSDS-UserPasswordExpiryTimeComputed")
	If @error = 0 Then
		If UBound($aTemp, 1) > 1 Then
			$aPwdInfo[13] = $aTemp[1][1]
		Else ; Required if DC version < 2008.
			$aPwdInfo[13] = $aPwdInfo[9]
			$iExtended = BitOR($iExtended, 8)
		EndIf
	Else
		$aPwdInfo[13] = $aPwdInfo[9]
		$iExtended = BitOR($iExtended, 16)
	EndIf
	Return SetError(0, $iExtended, $aPwdInfo)

EndFunc   ;==>_AD_GetPasswordInfo

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __AD_Int8ToSec
; Description ...: Function to convert Integer8 attributes from 64-bit numbers to seconds.
; Syntax.........: __AD_Int8ToSec($oInt8)
; Parameters ....: $oInt8 - 64-bit number (8 bytes) representing time in 100-nanosecond intervals.
; Return values .: Integer (Double Word) value
; Author ........: Jerold Schulman
; Modified.......: water
; Remarks .......: This function is used internally.
;                  Many attributes in Active Directory have a data type (syntax) called Integer8.
;                  These 64-bit numbers (8 bytes) often represent time in 100-nanosecond intervals.
; Related .......:
; Link ..........: http://www.rlmueller.net/Integer8Attributes.htm, http://msdn.microsoft.com/en-us/library/cc208659.aspx
; Example .......:
; ===============================================================================================================================
Func __AD_Int8ToSecEX($oInt8)

	Local $lngHigh, $lngLow
	$lngHigh = $oInt8.HighPart
	$lngLow = $oInt8.LowPart
	ConsoleWrite("Highpart: " & $lngHigh & " (dec), " & Hex($lngHigh) & " (hex); Lowpart: "  & $lngLow & " (dec), " & Hex($lngLow) & " (hex)" & @CRLF)
	If $lngLow < 0 Then $lngHigh = $lngHigh + 1
	Return -($lngHigh * (2 ^ 32) + $lngLow) / (10000000)

EndFunc   ;==>__AD_Int8ToSec