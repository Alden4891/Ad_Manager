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

if MsgBox(68,"", "Continue?") = 6 Then

;OPEN AD CONNECTION
_AD_Open("", "", "", "12-DC-02.ENTDSWD.LOCAL")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Open: An error has occurred.  @error: " & @error & ", @extended: " & @extended)
Local $sWorkbook = @ScriptDir & "\ad_create_template.xlsx"
Local $oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, $bReadOnly, $bVisible)
Local $aResult = _Excel_RangeRead($oWorkbook, Default, $oWorkbook.ActiveSheet.Usedrange.Columns("A:D"), 1)
Sleep(3000)
   For $i = 1 To UBound($aResult) - 1
	  $sUSERNAME =  $aResult[$i][0]
	  $sPASSWORD =  $aResult[$i][1]
	  $sCN		 =  $aResult[$i][2]
	  $sOU		 =  $aResult[$i][3]

	  ;ConsoleWrite(@CRLF & $sUSERNAME & ": Password Changed to " & $sPASSWORD)

	  if _AD_ObjectExists($sUSERNAME) Then
		 $aa =  _AD_GetObjectsInOU($sOU)
		 logger($sUSERNAME & " | | exists!")
	  else
		 logger($sUSERNAME & " | |ser not exists!")
	  EndIf
   Next

_AD_Close()

EndIf

Func logger($msg)
   _FileWriteLog(@ScriptDir & "\log\ad_acct_existence_check_" & @YEAR & "_" & @MON & "_" & @MDAY &  ".log", $msg)
   ConsoleWrite(@CRLF & $msg)
EndFunc

