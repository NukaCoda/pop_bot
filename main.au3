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
HotKeySet("+!d", "SearchColor") ; Shift-Alt-d
HotKeySet("+!s", "SetColorBox") ; Shift-Alt-s
HotKeySet("+!b", "SetColorBottom") ; Shift-Alt-b
HotKeySet("!1", "SetBoxTL") ; Shift-Alt-7
HotKeySet("!2", "SetBoxBR") ; Shift-Alt-3
HotKeySet("!3", "SetBottomTL") ; Shift-Ctrl-7
HotKeySet("!4", "SetBottomBR") ; Shift-Ctrl-3
HotKeySet("!5", "SetNextTL") ; Ctrl-Alt-7
HotKeySet("!6", "SetNextBR") ; Ctrl-Alt-3
HotKeySet("!7", "SetNextPos") ; Shift-Ctrl-2
HotKeySet("!8", "SetPickPos") ; Shift-Ctrl-1

While 1
	ToolTip("0x" & Hex($iColorBox, 6))
	Sleep(100)
WEnd

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

Func SearchColor()
	$iColorBox = InputBox("Color", "Enter color (0xRRGGBB) you wish to find." & @CR & " eg. For red enter 0xFF0000", "0xFF0000")
EndFunc   ;==>SearchColor

Func SetColorBox()
	$mousePos = MouseGetPos()
	$iColorBox = PixelGetColor($mousePos[0], $mousePos[1])
	MsgBox($MB_SYSTEMMODAL, "Title", "Searching for color code 0x" & Hex($iColorBox, 6) & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetColorBox

Func SetColorBottom()
	$mousePos = MouseGetPos()
	$iColorBottom = PixelGetColor($mousePos[0], $mousePos[1])
	$iTimeout = 5
	MsgBox($MB_SYSTEMMODAL, "Title", "Bottom of case set to color code 0x" & Hex($iColorBottom, 6) & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetColorBottom

Func SetBoxTL()
	$boxTL = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Title", "Box position Top Left set to x = " & $boxTL[0] & ", y = " & $boxTL[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetBoxTL

Func SetBoxBR()
	$boxBR = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Title", "Box position Bottom Right set to x = " & $boxBR[0] & ", y = " & $boxBR[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetBoxBR

Func SetNextTL()
	$nextTL = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Title", "Next position Top Left set to x = " & $nextTL[0] & ", y = " & $nextTL[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetNextTL

Func SetNextBR()
	$nextBR = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Title", "Next position Bottom Right set to x = " & $nextBR[0] & ", y = " & $nextBR[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetNextBR

Func SetBottomTL()
	$bottomTL = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Title", "Bottom of case position Top Left set to x = " & $bottomTL[0] & ", y = " & $bottomTL[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetBottomTL

Func SetBottomBR()
	$bottomBR = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Title", "Bottom of case position Bottom Right set to x = " & $bottomBR[0] & ", y = " & $bottomBR[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetBottomBR

Func SetPickPos()
	$pickButton = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Title", "Pick button position set to x = " & $pickButton[0] & ", y = " & $pickButton[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetPickPos

Func SetNextPos()
	$nextButton = MouseGetPos()
	MsgBox($MB_SYSTEMMODAL, "Title", "Next button position set to x = " & $nextButton[0] & ", y = " & $nextButton[1] & ". This message box will timeout after " & $iTimeout & " seconds or select the OK button.", $iTimeout)
EndFunc   ;==>SetNextPos