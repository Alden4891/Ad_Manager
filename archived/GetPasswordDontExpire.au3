#include "AD.au3"
#include "xls.au3"

_AD_Open("", "", "", "12-DC-02.ENTDSWD.LOCAL")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Open: An error has occurred.  @error: " & @error & ", @extended: " & @extended)

;get paswird that dont expired
Global $aResult1 = _AD_GetPasswordDontExpire("OU=Pantawid Pamilya,OU=Poverty Reduction Programs,OU=Operations and Programs Division,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL")
Global $aResult2 = _AD_GetPasswordDontExpire("OU=PPIS ML,OU=Clients,OU=FO12,DC=ENTDSWD,DC=LOCAL")
_ArrayConcatenate($aResult1,$aResult2)


_ArrayDisplay($aResult1)


_AD_Close()

_ArrayToXLS($aResult1, @ScriptDir & '\AD_PASSWORD_NEVER_EXPIRE_ACCOUNTS.xls')