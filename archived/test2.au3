#include "AD.au3"
#include "xls.au3"

_AD_Open("", "", "", "12-DC-02.ENTDSWD.LOCAL")
If  @error <> 0 Then Exit MsgBox(16, "LDAP", "Open: An error has occurred.  @error: " & @error & ", @extended: " & @extended)


Global $result = _AD_GetPasswordInfo(@UserName)
_ArrayDisplay($result)


_AD_Close()

;_ArrayToXLS($aResult, @ScriptDir & '\test1.xls')