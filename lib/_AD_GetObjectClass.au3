#AutoIt3Wrapper_AU3Check_Parameters= -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
; *****************************************************************************
; Example 1
; Displays the class for the logged on user and local computer.
; *****************************************************************************
#include <AD.au3>

; Open Connection to the Active Directory
_AD_Open()
If @error Then Exit MsgBox(16, "Active Directory Example Script", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)

; Return Main class / Structural class for the logged on user and local computer
MsgBox(64, "_AD_GetObjectClass - Example 1 - Display main class", _
		"Main class / Structural class for logged on user: " & _AD_GetObjectClass(@UserName) & @CRLF & _
		"Main class / Structural class for local computer: " & _AD_GetObjectClass(@ComputerName & "$"))

; Return Main class / Structural class plus superior classes from which the main class is deduced hierarchically for the logged on user
_ArrayDisplay(_AD_GetObjectClass(@UserName, True), "_AD_GetObjectClass - Example 2 - Display all classes for the logged on user")

; Close Connection to the Active Directory
_AD_Close()