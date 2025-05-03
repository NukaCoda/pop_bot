#include <WinAPIGdi.au3>
#include <MsgBoxConstants.au3>

; Press Esc key to terminate script.
; Press Home key to toggle search for color.
; Press Shift-Alt-d keys to enter a new color.
; Press Shift-Alt-s keys to color pick current mouse position.

Global $g_bPaused = False, $iColorBox = 0xFF00FF, $iColorBottom = 0xFF00FF, $iColorNext = 0x000000, $mousePos = [0,0]
Global $boxTL = [0,0], $boxBR = [0,0], $nextTL = [0,0], $nextBR = [0,0], $bottomTL = [0,0], $bottomBR = [0,0], $pickButton = [0,0], $nextButton = [0,0]
Global $iTimeout = 10
Global $scale = _WinAPI_EnumDisplaySettings('', $ENUM_CURRENT_SETTINGS)[0] / @DesktopWidth

HotKeySet("{HOME}", "TogglePause")
HotKeySet("{ESC}", "Terminate")
HotKeySet("+!b", "SetColors") ; Shift-Alt-b
HotKeySet("^!b", "SetColors") ; Ctrl-Alt-b
HotKeySet("!1", "SetPositions") ; Alt-1
HotKeySet("!2", "SetPositions") ; Alt-2
HotKeySet("!3", "SetPositions") ; Alt-3
HotKeySet("!4", "SetPositions") ; Alt-4
HotKeySet("!5", "SetPositions") ; Alt-5
HotKeySet("!6", "SetPositions") ; Alt-6
HotKeySet("!7", "SetPositions") ; Alt-7
HotKeySet("!8", "SetPositions") ; Alt-8

$ready = MsgBox($MB_OKCANCEL, "Pop Now Bot", "This program will automatically click through pages of Pop Now figures and click the 'Pick One to Shake' button when one is available. There are some steps to ensure the program runs smoothly. Ready to set up the parameters for Pop Now Bot?")

;MsgBox($MB_SYSTEMMODAL, "Title", "Previous box answer is " & $ready)

If ($ready = 1) Then
	Local $continue = 10
	While ($continue = 10)
		If (MsgBox($MB_OKCANCEL, "Box Locations", "Position your cursor at the upper left bounds of where the Pop Now boxes are located on your screen. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
			Terminate()
		EndIF
		$boxTL = MouseGetPos()
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Box Locations", "Top left boxes position set to x = " & $boxTL[0] & ", y = " & $boxTL[1] & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
	$continue = 10
	While ($continue = 10)
		If (MsgBox($MB_OKCANCEL, "Box Locations", "Position your cursor at the bottom right bounds of where the Pop Now boxes are located on your screen. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
			Terminate()
		EndIf
		$boxBR = MouseGetPos()
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Box Locations", "Bottom right boxes position set to x = " & $boxBR[0] & ", y = " & $boxBR[1] & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
	$continue = 10
	While ($continue = 10)
		MsgBox($MB_OKCANCEL, "Case Locations", "Position your cursor slight above and to the left of the bottom most portion of the Pop now case. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout)
		$bottomTL = MouseGetPos()
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Case Locations", "Top left of bottom of case position set to x = " & $bottomTL[0] & ", y = " & $bottomTL[1] & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
	$continue = 10
	While ($continue = 10)
		MsgBox($MB_OKCANCEL, "Case Locations", "Position your cursor slight below and to the right of the bottom most portion of the Pop now case. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout)
		$bottomBR = MouseGetPos()
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Case Locations", "Bottom right of bottom of case position set to x = " & $bottomBR[0] & ", y = " & $bottomBR[1] & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
	$continue = 10
	While ($continue = 10)
		MsgBox($MB_OKCANCEL, "Next Case Checker", "Position your cursor slight above and to the left of the POINT of the 'Next Case' button. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout)
		$nextTL = MouseGetPos()
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Next Case Checker", "Top left of next button point set to x = " & $nextTL[0] & ", y = " & $nextTL[1] & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
	$continue = 10
	While ($continue = 10)
		MsgBox($MB_OKCANCEL, "Next Case Checker", "Position your cursor slight below and to the right of the POINT of the 'Next Case' button. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout)
		$nextBR = MouseGetPos()
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Next Case Checker", "Bottom right of next button point set to x = " & $nextBR[0] & ", y = " & $nextBR[1] & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
	$continue = 10
	While ($continue = 10)
		MsgBox($MB_OKCANCEL, "Button Positions", "Position your cursor in the bottom quarter of the 'Next Case' button. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout)
		$nextButton = MouseGetPos()
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Button Positions", "Position of 'Next Case' button set to x = " & $nextButton[0] & ", y = " & $nextButton[1] & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
	$continue = 10
	While ($continue = 10)
		MsgBox($MB_OKCANCEL, "Button Positions", "Position your cursor on the 'Pick One to Shake' button. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout)
		$pickButton = MouseGetPos()
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Button Positions", "Position of 'Pick One to Shake' button set to x = " & $pickButton[0] & ", y = " & $pickButton[1] & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
	$continue = 10
	While ($continue = 10)
		MsgBox($MB_OKCANCEL, "Color Selection", "Position your cursor on a UNIQUE color of a Pop Now figure. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout)
		$mousePos = MouseGetPos()
		$iColorBox = PixelGetColor($mousePos[0], $mousePos[1])
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Color Selection", "Hex color code for Pop now box color is 0x" & Hex($iColorBox, 6) & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
	$continue = 10
	While ($continue = 10)
		MsgBox($MB_OKCANCEL, "Color Selection", "Position your cursor on a UNIQUE color of the very bottom of the Pop now case. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout)
		$mousePos = MouseGetPos()
		$iColorBottom = PixelGetColor($mousePos[0], $mousePos[1])
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Color Selection", "Hex color code for the bottom of the Pop now case color is 0x" & Hex($iColorBottom, 6) & ". Continue?")
	WEnd
	If ($continue = 2) Then
		Terminate()
	EndIf
Else
	Terminate()
EndIf

MsgBox($MB_OKCANCEL, "Pop Now Bot", "Set up complete! Press the 'Home' key to start and stop the program. Press the 'Escape' key to close or kill the program. Happy hunting!")

While 1
	ToolTip("Pop Now Bot | Escape to quit")
	Sleep(100)
WEnd
;Func SetColors()
;	Switch @HotKeyPressed
;		$mousePos = MouseGetPos()
;		Case "+!b"
;			$iColorBox = PixelGetColor($mousePos[0], $mousePos[1])
;		Case "^!b"
;			$iColorBottom = PixelGetColor($mousePos[0], $mousePos[1])
;	EndSwitch
;EndFunc

;Func SetPositions()
;	Switch @HotKeyPressed
;		Case "!1"
;			$boxTL = MouseGetPos()
;		Case "!2"
;			$boxBR = MouseGetPos()
;		Case "!3"
;			$bottomTL = MouseGetPos()
;		Case "!4"
;			$bottomBR = MouseGetPos()
;		Case "!5"
;			$nextTL = MouseGetPos()
;		Case "!6"
;			$nextBR = MouseGetPos()
;		Case "!7"
;			$pickButton = MouseGetPos()
;		Case "!8"
;			$nextButton = MouseGetPos()
;	EndSwitch
;EndFunc
	
Func TogglePause()
	$g_bPaused = Not $g_bPaused
	While $g_bPaused
		;Sleep(1150)
		$bottom = PixelSearch($bottomTL[0], $bottomTL[1], $bottomBR[0], $bottomBR[1], $iColorBottom, 0) ; Checks that the set has reached the bottom
		If IsArray($bottom) Then
			$boxes = PixelSearch($boxTL[0], $boxTL[1], $boxBR[0], $boxBR[1], $iColorBox, 0) ; Box location on screen
			$next = PixelSearch($nextTL[0], $nextTL[1], $nextBR[0], $nextBR[1], $iColorNext, 100) ; Next button
			#cs
				The search directions (from AutoIt help file - PixelSearch() function):
				Left-to-Right because "left" parameter < "right" parameter
				and
				Top-to-Bottom because "top" parameter < "bottom" parameter
				So, the top left search color is found first.
			#ce

			If IsArray($next) Then ; Next button IS loaded
				If IsArray($boxes) Then ; Box IS available
					MouseMove($pickButton[0], $pickButton[1], 0) ; Pick button
					MouseClick($MOUSE_CLICK_LEFT)
					$boxes = 0
					$g_bPaused = 0
				Else
					MouseMove($nextButton[0], $nextButton[1], 0) ; Next button
					ToolTip("READY")
					MouseClick($MOUSE_CLICK_LEFT)
					$boxes = 0
				EndIf
				$next = 0
			Else
				ToolTip("NOT READY") ; Next button loading
				$next = 0
			EndIf
		Else
		EndIf
	WEnd
	ToolTip("")
EndFunc   ;==>TogglePause

Func Terminate()
	Exit
EndFunc   ;==>Terminate

;Func SearchColor()
;	$iColorBox = InputBox("Color", "Enter color (0xRRGGBB) you wish to find." & @CR & " eg. For red enter 0xFF0000", "0xFF0000")
;EndFunc   ;==>SearchColor

;Func SetColorBox()
;	$mousePos = MouseGetPos()
;	$iColorBox = PixelGetColor($mousePos[0], $mousePos[1])
;	MsgBox($MB_SYSTEMMODAL, "Title", "Searching for color code 0x" & Hex($iColorBox, 6) & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetColorBox
;
;Func SetColorBottom()
;	$mousePos = MouseGetPos()
;	$iColorBottom = PixelGetColor($mousePos[0], $mousePos[1])
;	MsgBox($MB_SYSTEMMODAL, "Title", "Bottom of case set to color code 0x" & Hex($iColorBottom, 6) & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetColorBottom
;
;Func SetBoxTL()
;	$boxTL = MouseGetPos()
;	MsgBox($MB_SYSTEMMODAL, "Title", "Box position Top Left set to x = " & $boxTL[0] & ", y = " & $boxTL[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetBoxTL
;
;Func SetBoxBR()
;	$boxBR = MouseGetPos()
;	MsgBox($MB_SYSTEMMODAL, "Title", "Box position Bottom Right set to x = " & $boxBR[0] & ", y = " & $boxBR[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetBoxBR
;
;Func SetNextTL()
;	$nextTL = MouseGetPos()
;	MsgBox($MB_SYSTEMMODAL, "Title", "Next position Top Left set to x = " & $nextTL[0] & ", y = " & $nextTL[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetNextTL
;
;Func SetNextBR()
;	$nextBR = MouseGetPos()
;	MsgBox($MB_SYSTEMMODAL, "Title", "Next position Bottom Right set to x = " & $nextBR[0] & ", y = " & $nextBR[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetNextBR
;
;Func SetBottomTL()
;	$bottomTL = MouseGetPos()
;	MsgBox($MB_SYSTEMMODAL, "Title", "Bottom of case position Top Left set to x = " & $bottomTL[0] & ", y = " & $bottomTL[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetBottomTL
;
;Func SetBottomBR()
;	$bottomBR = MouseGetPos()
;	MsgBox($MB_SYSTEMMODAL, "Title", "Bottom of case position Bottom Right set to x = " & $bottomBR[0] & ", y = " & $bottomBR[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetBottomBR
;
;Func SetPickPos()
;	$pickButton = MouseGetPos()
;	MsgBox($MB_SYSTEMMODAL, "Title", "Pick button position set to x = " & $pickButton[0] & ", y = " & $pickButton[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetPickPos
;
;Func SetNextPos()
;	$nextButton = MouseGetPos()
;	MsgBox($MB_SYSTEMMODAL, "Title", "Next button position set to x = " & $nextButton[0] & ", y = " & $nextButton[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
;EndFunc   ;==>SetNextPos