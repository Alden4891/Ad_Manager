#include <Array.au3>
#include <Excel.au3>
#include "lib\AD.au3"
#include <File.au3>
#include <Date.au3>

; Create application object


; ****************************************************************************
; Open an existing workbook and return its object identifier.
; *****************************************************************************



;OPEN AD CONNECTION
_AD_Open("", "", "", "12-DC-02.ENTDSWD.LOCAL")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Open: An error has occurred.  @error: " & @error & ", @extended: " & @extended)


$sUSERNAME = "KAMAROHOM"

ConsoleWrite("Bulk Password Reset - "& @YEAR & "-" & @MON & "-" & @MDAY)
;~ _AD_EnableObject($sUSERNAME)
;~ ConsoleWrite(@CRLF & "_AD_EnableObject=" & @error)
;~ _AD_SetPassword($sUSERNAME,"WWEU9590",0)
;~ ConsoleWrite(@CRLF & "_AD_SetPassword=" & @error)
;~ _AD_DisablePasswordExpire($sUSERNAME)
;~ ConsoleWrite(@CRLF & "_AD_DisablePasswordExpire=" & @error)
;~ _AD_DisablePasswordChange ($sUSERNAME)
;~ ConsoleWrite(@CRLF & "_AD_DisablePasswordChange=" & @error)





_AD_Close()

