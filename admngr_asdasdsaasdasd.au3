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



;OPEN AD CONNECTION
_AD_Open("", "", "", "12-DC-02.ENTDSWD.LOCAL")



;_AD_AddUserToGroup("CN=SG-ALL-Employees,OU=Security Groups,OU=GROUPS,DC=OFFICE,DC=ORG","CN=Klaas Vaak (BDF-R),OU=NL,OU=Users,DC=OFFICE,DC=org")
$b = _AD_AddUserToGroup("CN=12_SSLVPN_GRP_ACL,OU=Groups,OU=FO12,DC=ENTDSWD,DC=LOCAL", "CN=AIZA T SULTAN,OU=PPIS ML,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL")
$c = _AD_AddUserToGroup("CN=12-SSL-groups,OU=Groups,OU=FO12,DC=ENTDSWD,DC=LOCAL", "CN=AIZA T SULTAN,OU=PPIS ML,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL")
$d = _AD_AddUserToGroup("CN=12-PPPP-GRP,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL", "CN=AIZA T SULTAN,OU=PPIS ML,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL")

ConsoleWrite(@CRLF & $a & ": " & @error)

;_AD_AddUserToGroup("12_SSLVPN_GRP_ACL", "ATSULTAN2")
;_AD_AddUserToGroup("12-PPPP-GRP", "ATSULTAN2")
;_AD_AddUserToGroup("12-SSL-groups", "ATSULTAN2")

_AD_Close()



Func logger($msg)
   _FileWriteLog(@ScriptDir & "\log\ad_create_acct_" & @YEAR & "_" & @MON & "_" & @MDAY &  ".log", $msg)
   ConsoleWrite(@CRLF & $msg)
EndFunc

