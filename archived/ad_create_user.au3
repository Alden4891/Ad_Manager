#include "AD.au3"
#include "xls.au3"
#include <Array.au3>

;OPEN AD CONNECTION
_AD_Open("", "", "", "12-DC-02.ENTDSWD.LOCAL")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Open: An error has occurred.  @error: " & @error & ", @extended: " & @extended)

$res = _AD_CreateUser($sOU, $sUser, $sCN, $bDynamic = False)
$res = _AD_SetPassword("NATUDON","TJNL7286",0)

   _AD_DisablePasswordExpire($sUser)
   _AD_DisablePasswordChange ($sUser)
   _AD_SetPassword($sUser, $randomStr,0)
ConsoleWrite($res)

_AD_Close()

