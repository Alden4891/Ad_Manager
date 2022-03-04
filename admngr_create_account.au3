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

; ****************************************************************************
; Open an existing workbook and return its object identifier.
; *****************************************************************************

if MsgBox(68,"", "You are about to bulk create account! Continue?") = 6 Then

;OPEN AD CONNECTION
_AD_Open("", "", "", "12-DC-02.ENTDSWD.LOCAL")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Open: An error has occurred.  @error: " & @error & ", @extended: " & @extended)
Local $sWorkbook = @ScriptDir & "\ad_create_template.xlsx"
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, $bReadOnly, $bVisible)
Local $aResult = _Excel_RangeRead($oWorkbook, "Template", $oWorkbook.ActiveSheet.Usedrange.Columns("A:D"), 1)
Sleep(3000)
   For $i = 1 To UBound($aResult) - 1
	  $sUSERNAME =  $aResult[$i][0]
	  $sPASSWORD =  $aResult[$i][1]
	  $sCN		 =  $aResult[$i][2]
	  $sOU		 =  $aResult[$i][3]

	  ;ConsoleWrite(@CRLF & $sUSERNAME & ": Password Changed to " & $sPASSWORD)

	  if _AD_ObjectExists($sUSERNAME) Then
		 logger($sUSERNAME & " exists!")
	  else
		 ;CREATE ACCOUNT
		 if _AD_CreateUser($sOU, $sUSERNAME, $sCN, False) then
			;SET MEMBERS
			if Ad_Set_members($sCN,$sOU) Then
			   ;SET PASSWORD
			   if _AD_SetPassword($sUSERNAME,$sPASSWORD,0) Then
				  _AD_ModifyAttribute($sUSERNAME, "Description", "Bulk Create User - " & _NowDate() & _NowTime())
				  _AD_DisablePasswordExpire($sUSERNAME)
				  _AD_DisablePasswordChange ($sUSERNAME)
				  logger($sUSERNAME & ": Successfully Created!")
			   else
				  if @error = 0 then
					 logger("Password for "& $sUSERNAME &" is not expired. Set Password aborted")
				  ElseIf @error = 1 then
					 logger("Account " & $sUSERNAME & " could not be found!")
				  Else
					 logger($sUSERNAME & " Unkown Error: set password error! ->" & @error)
				  EndIf
			   EndIf
			EndIf
		 else
			logger($sUSERNAME & " exists!")
		 EndIf
	  EndIf
   Next

_AD_Close()

EndIf

Func Ad_Set_members($oCN, $oOU)

;~    Local $a = _AD_AddUserToGroup("CN=12_SSLVPN_GRP_ACL,OU=Groups,OU=FO12,DC=ENTDSWD,DC=LOCAL", "CN=" & $oCN & "," & $oOU)
;~    Local $b = _AD_AddUserToGroup("CN=12-SSL-groups,OU=Groups,OU=FO12,DC=ENTDSWD,DC=LOCAL", "CN=" & $oCN & "," & $oOU)
;~    Local $c = _AD_AddUserToGroup("CN=12-PPPP-GRP,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL", "CN=" & $oCN & "," & $oOU)
   Local $d = _AD_AddUserToGroup("CN=PPIS-ML-12,OU=PPIS ML,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL", "CN=" & $oCN & "," & $oOU)
   Local $ret = 0
   if $d = 1 Then
	  logger($oCN & " successfully added to 12_SSLVPN_GRP_ACL,12-SSL-groups, and 12-PPPP-GRP groups!   ")
	  $ret = 0
   elseif @error = 1 then
	  logger("Target groups do not exist!")
   elseif @error = 2 then
	  logger("User, group or computer to be added to is not found!")
   elseif @error = 3 then
	  logger("Already a member of the group!")
   elseif @error = -2147352567 then
	  logger("Add Group denied! You dont have sufficient permission.")
   else
	  logger("Unknown error!" & $a & " " & $d)
   EndIf
EndFunc

Func logger($msg)
   _FileWriteLog(@ScriptDir & "\log\ad_create_acct_" & @YEAR & "_" & @MON & "_" & @MDAY &  ".log", $msg)
   ConsoleWrite(@CRLF & $msg)
EndFunc

