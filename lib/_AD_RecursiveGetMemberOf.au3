#AutoIt3Wrapper_AU3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
; *****************************************************************************
; Example 1
; Returns a recursively searched list of groups the currently logged on user
; is a member of.
; For groups that are inherited, the FQDN of the group or user, and the FQDN(s)
; of the group(s) it was inherited from, seperated by '|'
; *****************************************************************************
#include <AD.au3>

; Open Connection to the Active Directory
_AD_Open()
If @error Then Exit MsgBox(16, "Active Directory Example Skript", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)

; Returns a recursively searched list of groups the currently logged on user is a member of
Global $aUser = _AD_RecursiveGetMemberOf(@UserName, 10, True)
If @error > 0 Then
	MsgBox(64, "Active Directory Functions - Example 1", "User '" & @UserName & "' has not been assigned to any group")
Else
	; For groups that are inherited, the return is the FQDN of the group or user, and the FQDN(s) of the group(s) it
	; was inherited from, seperated by '|'
	_ArrayDisplay($aUser, "Active Directory Functions - Example 1 - Group names user '" & @UserName & "' is a member of")
EndIf

; Close Connection to the Active Directory
_AD_Close()