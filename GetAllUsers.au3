#include "lib\AD.au3"
#include "xls.au3"
#include <Array.au3>

;OPEN AD CONNECTION
_AD_Open("", "", "", "12-DC-02.ENTDSWD.LOCAL")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Open: An error has occurred.  @error: " & @error & ", @extended: " & @extended)

;GET PANTAWID AD ACCOUNTS
$aResult1 = _AD_GetObjectsInOU("OU=Pantawid Pamilya,OU=Poverty Reduction Programs,OU=Operations and Programs Division,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL", "(&(objectclass=user)(!(employeenumber=stage)))", 2, "userPrincipalName,sAMAccountName,distinguishedName,cn,company,displayname,description,mail,co")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Query: An error has  occurred.  @error: " & @error & ", @extended: " & @extended)

;GET ML AD ACCOUNT
$aResult2 = _AD_GetObjectsInOU("OU=PPIS ML,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL", "(&(objectclass=user)(!(employeenumber=stage)))", 2, "userPrincipalName,sAMAccountName,distinguishedName,cn,company,displayname,description,mail,co")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Query: An error has  occurred.  @error: " & @error & ", @extended: " & @extended)

_ArrayConcatenate($aResult1,$aResult2)
$upper_bound =  UBound($aResult1);

for $i = 1 to $upper_bound - 1

   $userPrincipalName = $aResult1[$i][0]
   $sAMAccountName    = $aResult1[$i][1]
   $distinguishedName = $aResult1[$i][2]
   $cn                = $aResult1[$i][3]
   $company           = $aResult1[$i][4]
   $displayname       = $aResult1[$i][5]
   $description       = $aResult1[$i][6]
   $mail              = $aResult1[$i][7]
;   $co1                = $aResult1[$i][8]
;   $co2                = $aResult1[$i][9]

   Global $pwInfo = _AD_GetPasswordInfo($sAMAccountName)
   ;ConsoleWrite(@CRLF & "->" & UBound($pwInfo) & " ")
   if UBound($pwInfo) = 14 Then
	  $last_changed = $pwInfo[8]
	  $last_change = $pwInfo[9]
	  $aResult1[$i][8] = $pwInfo[8]  & " | " & $pwInfo[9] & "|" & _AD_IsPasswordExpired($sAMAccountName) & "|" & _AD_IsAccountExpired($sAMAccountName);last changed & ;Password Expiration

	  ConsoleWrite(@CRLF & $userPrincipalName & " - Password Last Changed: " & $pwInfo[8] & " and " & "Password Expires on: " & $pwInfo[9] & " [" & ($i/$upper_bound)*100 & "%]")
   Else
	  $aResult1[$i][8] = "ERROR!"
   EndIf
Next

; _ArrayDisplay($aResult1)

;ConsoleWrite(@CRLF & $sAMAccountName)
_AD_Close()

_ArrayToXLS($aResult1, @ScriptDir & '\AD_ALLUSERS.xls')