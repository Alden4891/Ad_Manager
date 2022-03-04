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
Local $R = "A" & $var & ":" & "B" & $var

; ****************************************************************************
; Open an existing workbook and return its object identifier.
; *****************************************************************************

if MsgBox(68,"", "You are about to bulk reset password! Continue?") = 6 Then

;OPEN AD CONNECTION
_AD_Open("", "", "", "12-DC-02.ENTDSWD.LOCAL")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Open: An error has occurred.  @error: " & @error & ", @extended: " & @extended)


Local $sWorkbook = @ScriptDir & "\ad_reset_template.xlsx"
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, $bReadOnly, $bVisible)
Local $aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.ActiveSheet.Usedrange.Columns("A:B"), 1)
Sleep(5000)
   For $i = 1 To UBound($aResult) - 1
	  $sUSERNAME =  $aResult[$i][0]
	  $sPASSWORD =  $aResult[$i][1]
	  ;ConsoleWrite(@CRLF & $sUSERNAME & ": Password Changed to " & $sPASSWORD)

	  ;FORCE ENABLE OBJECT
	  _AD_EnableObject($sUSERNAME)
	  ConsoleWrite(@CRLF & "_AD_EnableObject=" & @error)
	  $e1 = @error
	  _AD_SetPassword($sUSERNAME,$sPASSWORD,0)
	  ConsoleWrite(@CRLF & "_AD_SetPassword=" & @error)
	  $e2 = @error
	  _AD_DisablePasswordExpire($sUSERNAME)
	  ConsoleWrite(@CRLF & "_AD_DisablePasswordExpire=" & @error)
	  $e3 = @error
	  _AD_DisablePasswordChange ($sUSERNAME)
	  ConsoleWrite(@CRLF & "_AD_DisablePasswordChange=" & @error)
	  $e4 = @error

	  if ($e1+$e2+$e3+$e4) = 0 Then
		 _FileWriteLog ( @ScriptDir & "\reset_file.log" , "|" & $sUSERNAME & "|ok")
		 ConsoleWrite(@CRLF & "|" & $sUSERNAME & "|ok")
		 _AD_ModifyAttribute($sUSERNAME, "Description", "Bulk Password Reset - "& @YEAR & "-" & @MON & "-" & @MDAY)
	  Else
		 _FileWriteLog ( @ScriptDir & "\reset_file.log" , "|" & $sUSERNAME & "|failed")
		 ConsoleWrite(@CRLF & "|" & $sUSERNAME & "|failed")
	  EndIf

;~ 	  if _AD_IsPasswordExpired($sUSERNAME) Then
;~ 		  ConsoleWrite($i & " ")
;~ 		 if _AD_SetPassword($sUSERNAME,$sPASSWORD,0) Then
;~ 			logger($sUSERNAME & ": Password Changed to " & $sPASSWORD)
;~ 			_AD_ModifyAttribute($sUSERNAME, "Description", "Bulk Password Reset - " & _NowTime())
;~ 			_AD_DisablePasswordExpire($sUSERNAME)
;~ 			_AD_DisablePasswordChange ($sUSERNAME)
;~ 		 else
;~ 			if @error = 0 then
;~ 			   logger("Password for "& $sUSERNAME &" has not expired. Reset aborted")
;~ 			ElseIf @error = 1 then
;~ 			   logger("Account " & $sUSERNAME & " could not be found!")
;~ 			Else
;~ 			   logger($sUSERNAME & " Unkown Error: Reset aborted! ->" & @error)
;~ 			EndIf
;~ 		 EndIf
;~ 	  else
;~ 		 logger($sUSERNAME & " is not expired! Reset aborted.")
;~ 	  EndIf
   Next

_AD_Close()

EndIf

Func logger($msg)
   _FileWriteLog(@ScriptDir & "\log\ad_pass_reset_" & @YEAR & "_" & @MON & "_" & @MDAY &  ".log", $msg)
   ConsoleWrite(@CRLF & $msg)
EndFunc

