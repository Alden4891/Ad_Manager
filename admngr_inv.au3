#include <Array.au3>
#include <Excel.au3>
#include "lib\AD.au3"
#include <File.au3>
#include <Date.au3>

; Create application object
Local $oExcel = _Excel_Open()
Local $bReadOnly = False
Local $bVisible = True
Local $var = 1
Local $R = "A" & $var & ":" & "D" & $var

Local $root1 = "OU=Pantawid Pamilya,OU=Poverty Reduction Programs,OU=Operations and Programs Division,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL"
Local $root2 = "OU=PPIS ML,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL"

; ****************************************************************************
; Open an existing workbook and return its object identifier.
; *****************************************************************************

;OPEN AD CONNECTION
_AD_Open("aaquinones", "protectme@4891021408271209", "", "12-DC-02.ENTDSWD.LOCAL")

If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Open: An error has occurred.  @error: " & @error & ", @extended: " & @extended)
Local $sWorkbook = @ScriptDir & "\PANTAWID_AD_RESULT.xlsx"
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, $bReadOnly, $bVisible)
Local $aResult = _Excel_RangeRead($oWorkbook, "AD_PROF", $oWorkbook.ActiveSheet.Usedrange.Columns("j:j"), 1)
Local $offset = 400 ;2=default

For $i = $offset To UBound($aResult) - 1
   Local $username = $aResult[$i]
   Local $perc = Round(($i/(UBound($aResult) - 1))*100,2) & "% complete"
   ToolTip("Progress [" & $username & "]: " & $perc,0,0)

   if Mod($i,15)=9 then
;~ 	  Send("{ctrldown}s{ctrlup}")
;~ 	  ConsoleWrite(@CRLF & "Saved!")
;~ 	  Sleep(5000)
   Else
	  ConsoleWrite(@CRLF & " [" & $i & "][" & $username & "] " & $perc)
   EndIf

   ;STATUS
   $ObjectDisabled = isObjectDisabled($username)
   _Excel_RangeWrite($oWorkbook, "AD_PROF",$ObjectDisabled,"m" & $i+1)
   if $ObjectDisabled <> "NF" then
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isPasswordExpired($username),"n" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isAccountExpired($username),"o" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isObjectLocked($username),"p" & $i+1)

;~ 	  ;RIGHTS
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",hasUnlockResetRights($username),"q" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",hasFullRights($username),"r" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",hasGroupUpdateRights($username),"s" & $i+1)

	  ;MEMBERSHIP
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12_DHCP_Administrators",$username),"t" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12_LyncDel_RoomAdmin",$username),"u" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12_OBTR_RDP_FTP_Administrator",$username),"v" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12_OU_Administrators",$username),"w" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12_SCCM_Administrators",$username),"x" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12_SSLVPN_GRP_ACL",$username),"y" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12_SVRGRP_Administrators",$username),"z" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12_User_Administrators",$username),"aa" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12-IT-wAdmin-Group",$username),"ab" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12-RICTMS-GRP",$username),"ac" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12-RICTMU-Group",$username),"ad" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12-SSL-groups",$username),"ae" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12-PPPP-GRP",$username),"af" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("PPIS-ML-12",$username),"ag" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",isMemberOf("12_PPISML_VPN_HTTP",$username),"ah" & $i+1)

;~ 	  ;EXPIRATION & LOCKOUTS
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",getPassExpiration($username),"ai" & $i+1)
	  _Excel_RangeWrite($oWorkbook, "AD_PROF",getPassLastChanged($username),"aj" & $i+1)
;~ 	  _Excel_RangeWrite($oWorkbook, "AD_PROF",getLastLoginDate($username),"ak" & $i+1)




   EndIf



Next
Func isMemberOf($group,$username)
   ; Get the Fully Qualified Domain Name (FQDN) for the current user
   $sFQDN_User = _AD_SamAccountNameToFQDN($username)

   ; Get an array of group names (FQDN) that the current user is immediately a member of
   $aUser = _AD_GetUserGroups($username)
   $ret_value = "";
   For $sFQDN_Group In $aUser
	  ; Check the group membership of the specified user for the specified group
	  if StringInStr($sFQDN_Group,$group) then
		 $ret_value = 1
		 ExitLoop
	  EndIf
;~ 	  $iResult = _AD_IsMemberOf($sFQDN_Group, $sFQDN_User)
   Next
   Return $ret_value
EndFunc

;RIGHTS UDF
Func hasGroupUpdateRights($username)
   if _AD_HasGroupUpdateRights($username,$username) = 1 Then
	  return 1
   Else
	  Switch @error
		 case 1 to 2
			return "NF"
		 case else
			Return ""
	  EndSwitch
   EndIf
EndFunc
Func hasFullRights($username)
   if _AD_HasFullRights($username,$username) = 1 Then
	  return 1
   Else
	  Switch @error
		 case 1 to 2
			return "NF"
		 case else
			Return ""
	  EndSwitch
   EndIf
EndFunc

Func hasUnlockResetRights($username)
   if _AD_HasUnlockResetRights($username) = 1 Then
	  return 1
   Else
	  Switch @error
		 case 1 to 2
			return "NF"
		 case else
			Return ""
	  EndSwitch
   EndIf
EndFunc

;STATUS UDF
Func getPassExpiration($username)
   $res = _AD_GetPasswordInfo($username)
   if @error > 0 Then
	  return ""
   else
	  return  $res[11]
   EndIf
EndFunc

;aj
Func getPassLastChanged($username)
   $res = _AD_GetPasswordInfo($username)
   if @error > 0 Then
	  return ""
   else
	  return  $res[10]
   EndIf
EndFunc

;ak
Func getLastLoginDate($username)
   $res = _AD_GetLastLoginDate($username)
   if @error > 0 Then
	  _ArrayDisplay($res)
	  return "Err: " & @error

   else
	  return  $res
   EndIf
EndFunc

Func isPasswordNeverExpire($username)
   $res = _AD_GetPasswordInfo($username)
   return  Abs(_DateIsValid($res[11])+-1)
EndFunc

Func isObjectDisabled($username)
   if _AD_IsObjectDisabled($username) = 1 Then
	  return 1
   Else
	  Switch @error
		 case 1
			return "NF"
		 case else
			Return ""
	  EndSwitch
   EndIf
EndFunc

Func isPasswordExpired($username)
   if _AD_IsPasswordExpired($username) = 1 Then
	  if isPasswordNeverExpire($username) Then
		 return "NE"
	  Else
		 return 1
	  EndIf
   Else
	  Switch @error
		 case 1
			return "NF"
		 case 0
			Return ""
		 case else
			Return "Error: " & @error
	  EndSwitch
   EndIf
EndFunc

Func isAccountExpired($username)
   if _AD_IsAccountExpired($username) = 1 Then
	  return 1
   Else
	  Switch @error
		 case 1 to 2
			return "NF"
		 case else
			Return ""
	  EndSwitch
   EndIf
EndFunc

Func isObjectLocked($username)
   if _AD_IsObjectLocked($username) = 1 Then
	  return 1
   Else
	  Switch @error
		 case 1
			return "NF"
		 case else
			Return ""
	  EndSwitch
   EndIf
EndFunc




_AD_Close()

;~ Func logger($msg)
;~    _FileWriteLog(@ScriptDir & "\log\ad_create_acct_" & @YEAR & "_" & @MON & "_" & @MDAY &  ".log", $msg)
;~    ConsoleWrite(@CRLF & $msg)
;~ EndFunc

