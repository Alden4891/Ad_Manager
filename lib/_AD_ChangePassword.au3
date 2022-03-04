#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_AU3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7 
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; *****************************************************************************
; Example 1
; Sets the password for the currently logged on user
; *****************************************************************************
#include <AD.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <MsgBoxConstants.au3>

Global $bShow = False

; Open Connection to the Active Directory
_AD_Open()
If @error Then Exit MsgBox(16, "Active Directory Example Skript", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)

Global $iReply = MsgBox(308, "Active Directory Functions - Example 1", "This script changes the password for the current user." & @CRLF & @CRLF & _
		"Are you sure you want to change the Active Directory?")
If $iReply <> 6 Then Exit

; Enter user and password to change
#region ### START Koda GUI section ### Form=
Global $Form1 = GUICreate("Active Directory Functions - Example 1", 414, 156, 251, 112)
GUICtrlCreateLabel("Old (current) password:", 8, 10, 231, 17)
GUICtrlCreateLabel("New password:", 8, 42, 121, 17)
GUICtrlCreateLabel("New password:", 8, 72, 121, 17)
Global $IOldPW = GUICtrlCreateInput("", 241, 8, 159, 21, $ES_PASSWORD)
Global $INewPW1 = GUICtrlCreateInput("", 241, 40, 159, 21, $ES_PASSWORD)
Global $INewPW2 = GUICtrlCreateInput("", 241, 72, 159, 21, $ES_PASSWORD)
Global $BOK = GUICtrlCreateButton("Change Password", 8, 104, 121, 33)
Global $BShowPW = GUICtrlCreateButton("Show Passwords", 200, 104, 110, 33)
Global $BCancel = GUICtrlCreateButton("Cancel", 328, 104, 73, 33, BitOR($GUI_SS_DEFAULT_BUTTON, $BS_DEFPUSHBUTTON))
GUISetState(@SW_SHOW)
#endregion ### END Koda GUI section ###

Global $sDefaultPassChar = GUICtrlSendMsg($IOldPW, $EM_GETPASSWORDCHAR, 0, 0)
While 1
	Global $nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $BCancel
			Exit
		Case $BShowPW
			If $bShow = False Then
				GUICtrlSendMsg($IOldPW, $EM_SETPASSWORDCHAR, 0, 0)
				GUICtrlSendMsg($INewPW1, $EM_SETPASSWORDCHAR, 0, 0)
				GUICtrlSendMsg($INewPW2, $EM_SETPASSWORDCHAR, 0, 0)
				GUICtrlSetData($BShowPW, "Hide Passwords")
				$bShow = True
			Else
				GUICtrlSendMsg($IOldPW, $EM_SETPASSWORDCHAR, $sDefaultPassChar, 0)
				GUICtrlSendMsg($INewPW1, $EM_SETPASSWORDCHAR, $sDefaultPassChar, 0)
				GUICtrlSendMsg($INewPW2, $EM_SETPASSWORDCHAR, $sDefaultPassChar, 0)
				GUICtrlSetData($BShowPW, "Show Passwords")
				$bShow = False
			EndIf
			GUICtrlSetState($IOldPW, $GUI_FOCUS) ;Input needs focus to redraw characters
			GUICtrlSetState($INewPW1, $GUI_FOCUS)
			GUICtrlSetState($INewPW2, $GUI_FOCUS)
		Case $BOK
			Global $sOldPW = GUICtrlRead($IOldPW)
			Global $sNewPW1 = GUICtrlRead($INewPW1)
			Global $sNewPW2 = GUICtrlRead($INewPW2)
			if $sNewPW1 = $sNewPW2 Then ExitLoop
			MsgBox($MB_ICONERROR, "Active Directory - Change Password", "New passwords don't match!")
		EndSwitch
WEnd

; Change the password
Global $iValue = _AD_ChangePassword($sOldPW, $sNewPW1)
If $iValue = 1 Then
	MsgBox(64, "Active Directory Functions - Example 1", "Password for the current user successfully changed")
ElseIf @error = 1 Then
	MsgBox(16, "Active Directory Functions - Example 1", "Error occurred when accessing the current user object!" & @LF & "@error = " & @error & ", @extended = " & @extended)
Else
	MsgBox(16, "Active Directory Functions - Example 1", "Error occurred when changing the password!" & @LF & "@error = " & @error & ", @extended = " & @extended)
EndIf

; Close Connection to the Active Directory
_AD_Close()