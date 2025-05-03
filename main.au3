;version 10
#include <WinAPIGdi.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>

Global $g_bPaused = False, $iColorBox = 0xFF00FF, $iColorBottom = 0xFF00FF, $iColorNext = 0x000000, $mousePos = [0,0]
Global $boxTL = [0,0], $boxBR = [0,0], $next = [0,0], $bottom = [0,0], $pickButton = [0,0], $nextButton = [0,0], $errorTL = [0,0], $errorBR = [0,0]
Global $iTimeout = 10
Global $scale = _WinAPI_EnumDisplaySettings('', $ENUM_CURRENT_SETTINGS)[0] / @DesktopWidth
Global $casesPassed = 0
Global $errorChecksum, $nextChecksum, $caseChecksum
Global $debugTime = 200

HotKeySet("{HOME}", "TogglePause")
HotKeySet("{ESC}", "Terminate")

$ready = MsgBox($MB_OKCANCEL, "Pop Now Bot", "This program will automatically click through pages of Pop Now figures and click the 'Pick One to Shake' button when one is available. There are some steps to ensure the program runs smoothly. Ready to set up the parameters for Pop Now Bot?")

If ($ready = 1) Then
	If (MsgBox($MB_YESNO, "Load Preset", "Do you want to load a preset?") = 7) Then
		Local $continue = 10
		While ($continue = 10) ; Set TL bounds for boxes
			If (MsgBox($MB_OKCANCEL, "Box Locations", "Position your cursor at the upper left bounds of where the Pop Now boxes are located on your screen. Position will be saved in " & $iTimeout & "seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIF
			$boxTL = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Box Locations", "Top left boxes position set to x = " & $boxTL[0] & ", y = " & $boxTL[1] & ". Continue?")
		WEnd
		If ($continue = 2) Then
			Terminate()
		EndIf

		$continue = 10
		While ($continue = 10) ; Set BR bounds for boxes
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
		While ($continue = 10) ; Set position of bottom of case
			If (MsgBox($MB_OKCANCEL, "Case Locations", "Position your cursor on a unique color on the bottom of the case. Position will be saved in " & $iTimeout & " seconds or click cancel.",$iTimeout) = 2) Then
				Terminate()
			EndIf
			$bottom = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Case Locations", "Position set to x = " & $bottom[0] & ", y = " & $bottom[1] & ". Continue?")
		WEnd
		If ($continue = 2) Then
			Terminate()
		EndIf

		$continue = 10
		While ($continue = 10) ; Set TL bounds for error message
			If (MsgBox($MB_OKCANCEL, "Error Locations", "Position your cursor to the top left position of the error message. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIf
			$errorTL = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Error Locations", "Top left of error message position set to x = " & $errorTL[0] & ", y = " & $errorTL[1] & ". Continue?")
		WEnd
		If ($continue = 2) Then
			Terminate()
		EndIf

		$continue = 10
		While ($continue = 10) ; Set BR bounds for error message
			If (MsgBox($MB_OKCANCEL, "Error Locations", "Position your cursor in the bottom right position of the error message. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIf
			$errorBR = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Error Locations", "Bottom right of error message position set to x = " & $errorBR[0] & ", y = " & $errorBR[1] & ". Continue?")
		WEnd
		If ($continue = 2) Then
			Terminate()
		EndIf

		$continue = 10
		While ($continue = 10) ; Set position for 'next' checker
			If (MsgBox($MB_OKCANCEL, "Next Case Checker", "Position your cursor in the center of the 'Next' button. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2)Then
				Terminate()
			EndIf
			$next = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Next Case Checker", "Position set to x = " & $next[0] & ", y = " & $next[1] & ". Continue?")
		WEnd
		If ($continue = 2) Then
			Terminate()
		EndIf

		$continue = 10
		While ($continue = 10) ; set position of 'next' button
			If (MsgBox($MB_OKCANCEL, "Button Positions", "Position your cursor in the bottom quarter of the 'Next Case' button. Position will be saved in " & $iTimeout & " seconds or click cancel.",$iTimeout) = 2) Then
				Terminate()
			EndIf
			$nextButton = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Button Positions", "Position of 'Next Case' button set to x = " & $nextButton[0] & ", y = " & $nextButton[1] & ". Continue?")
		WEnd
		If ($continue = 2) Then
			Terminate()
		EndIf

		$continue = 10
		While ($continue = 10) ; Set position of 'pick one' button
			If (MsgBox($MB_OKCANCEL, "Button Positions", "Position your cursor on the 'Pick One to Shake' button. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2)Then
				Terminate()
			EndIf
			$pickButton = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Button Positions", "Position of 'Pick One to Shake' button set to x = " & $pickButton[0] & ", y = " & $pickButton[1] & ". Continue?")
		WEnd
		If ($continue = 2) Then
			Terminate()
		EndIf

		$continue = 10
		While ($continue = 10) ; Set color to search for available box
			If (MsgBox($MB_OKCANCEL, "Color Selection", "Position your cursor on a UNIQUE color of a Pop Now figure. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIf
			$mousePos = MouseGetPos()
			$iColorBox = PixelGetColor($mousePos[0], $mousePos[1])
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Color Selection", "Hex color code for Pop Now box color is 0x" & Hex($iColorBox, 6) & ". Continue?")
		WEnd
		If ($continue = 2) Then
			Terminate()
		EndIf
	Else
		LoadPreset()
	EndIf
Else
	Terminate()
EndIf ;==== Finish setup

Sleep(500)

$mouseTemp = MouseGetPos()
MouseMove(0, 0, 0)

$errorChecksum = PixelChecksum($errorTL[0], $errorTL[1], $errorBR[0], $errorBR[1])
$nextChecksum = PixelChecksum($next[0] - 10, $next[1] - 10, $next[0] + 10, $next[1] + 10)
$caseChecksum = PixelChecksum($bottom[0] - 10, $bottom[1] - 10, $bottom[0] + 10, $bottom[1] + 10)

MouseMove($mouseTemp[0], $mouseTemp[1], 0)

Sleep(500)

;Save Preset ===================
If (MsgBox($MB_YESNO, "Save Preset", "Would you like to save these parameters as a preset?") = 6) Then
	SavePreset()
EndIf

MsgBox($MB_OK, "Pop Now Bot", "Set up complete! Press the 'Home' key to start and stop the program. Press the 'Escape' key to close or kill the program. Happy hunting!")

While 1
	ToolTip("Pop Now Bot | PAUSED | Home to resume, Escape to quit")
	Sleep(50)
WEnd
	
Func TogglePause()
	$g_bPaused = Not $g_bPaused
	While Not $g_bPaused
		ToolTip("Pop Now Bot | PAUSED | Home to resume, Escape to quit")
		Sleep(50)
	Wend
	While $g_bPaused
		ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
		;ToolTip("debug | next search") ;========================================
		;Sleep($debugTime) ;===========================================================

		If ($nextChecksum = PixelChecksum($next[0] - 10, $next[1] - 10, $next[0] + 10, $next[1] + 10)) Then ; Next button loaded
			ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
			;ToolTip("debug | next good") ;========================================
			;Sleep($debugTime) ;=========================================================
			Sleep(1000)
			If ($caseChecksum = PixelChecksum($bottom[0] - 10, $bottom[1] - 10, $bottom[0] + 10, $bottom[1] + 10)) Then
				ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
				ToolTip("debug | case good") ;========================================
				;Sleep($debugTime) ;=========================================================
				$boxes = PixelSearch($boxTL[0], $boxTL[1], $boxBR[0], $boxBR[1], $iColorBox, 0) ; Box location on screen
				;ToolTip("debug | box search") ;========================================
				;Sleep($debugTime) ;==========================================================

				If IsArray($boxes) Then ; Box is available
					ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
					;ToolTip("debug | box good") ;========================================
					;Sleep($debugTime) ;========================================================
					MouseMove($pickButton[0], $pickButton[1], 0) ; Pick button
					MouseClick($MOUSE_CLICK_LEFT)
					$casesPassed = 0
					$boxes = 0
					$errorPresent = 0
					$errorTimer = 0
					While ($errorTimer <= 100)
						ToolTip($errorTimer & "% Lie Detecting")
						Sleep(20)
						ToolTip($errorTimer & "% Lie Detecting")
						Sleep(20)
						ToolTip($errorTimer & "% Lie Detecting")
						If ($errorChecksum = PixelChecksum($errorTL[0], $errorTL[1], $errorBR[0], $errorBR[1])) Then
							$errorTimer = $errorTimer + 2
						Else
							$errorPresent = $errorPresent + 1
							$errorTimer = $errorTimer + 2
						EndIf
						$g_bPaused = 0
						Sleep(9)
					Wend

					If ($errorPresent > 0) Then
						ToolTip("LIE DETECTED - Resuming search")
						$g_bPaused = 1
						Sleep(5000)
						MouseMove($nextButton[0], $nextButton[1], 0) ; Next button ready
						ToolTip("LIE DETECTED - Resuming search")
						MouseClick($MOUSE_CLICK_LEFT)
						Sleep(2000)
						$boxes = 0
						ToolTip("LIE DETECTED - Resuming search")
					Else
						$g_bPaused = 0
						While Not $g_bPaused
							ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
							Sleep(50)
						Wend
					EndIf
				Else
					ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
					;ToolTip("debug | box bad") ;========================================
					;Sleep($debugTime) ;=======================================================
					MouseMove($nextButton[0], $nextButton[1], 0) ; Next button ready
					MouseClick($MOUSE_CLICK_LEFT)
					;Sleep(100)
					$boxes = 0
					$case = 0
					ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
				EndIf ;== Box check
				$loaded = 0
				ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
			Else
				ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
				;ToolTip("debug | case bad") ;========================================
				;Sleep($debugTime) ;=========================================================
				$loaded = 0
			EndIf ;== Case check
			ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
		Else
			ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
		EndIf ;== Next check
		ToolTip("Pop Now Bot | Home to Pause, Escape to quit")
	WEnd ;== While paused
	ToolTip("")
EndFunc ;== TogglePause

Func LoadPreset()
	$lPresetName = InputBox("Load Preset", "Presets location: C:\Pop Now Bot Presets" & @CRLF & @CRLF & "Enter the name of the preset.")

	$rFileName = "C:\Pop Now Bot Presets\" & $lPresetName & ".txt"

	$rFileHandle = FileOpen($rFileName, $FO_READ)

	If $rFileHandle = -1 Then
        MsgBox($MB_SYSTEMMODAL, "ERROR", "An error occurred when reading the file." & @CRLF & @CRLF & "Make sure the preset file is in the 'Pop Now Bot Preset' folder located on your C drive." & @CRLF & @CRLF & "Closing application")
		FileClose($rFileHandle)
        Terminate()
    EndIf

	$boxTL[0]		= FileReadLine($rFileHandle, 1)
	$boxTL[1]		= FileReadLine($rFileHandle, 2)
	$boxBR[0]		= FileReadLine($rFileHandle, 3)
	$boxBR[1]		= FileReadLine($rFileHandle, 4)
	$bottom[0]		= FileReadLine($rFileHandle, 5)
	$bottom[1]		= FileReadLine($rFileHandle, 6)
	$errorTL[0]		= FileReadLine($rFileHandle, 7)
	$errorTL[1]		= FileReadLine($rFileHandle, 8)
	$errorBR[0]		= FileReadLine($rFileHandle, 9)
	$errorBR[1]		= FileReadLine($rFileHandle, 10)
	$next[0]		= FileReadLine($rFileHandle, 11)
	$next[1]		= FileReadLine($rFileHandle, 12)
	$nextButton[0]	= FileReadLine($rFileHandle, 13)
	$nextButton[1]	= FileReadLine($rFileHandle, 14)
	$pickButton[0]	= FileReadLine($rFileHandle, 15)
	$pickButton[1]	= FileReadLine($rFileHandle, 16)
	$iColorBox		= FileReadLine($rFileHandle, 17)

	MsgBox($MB_SYSTEMMODAL, "Loaded Content", "Top Left of boxes [x] = " & $boxTL[0] & @CRLF & "Top Left of boxes [y] = " & $boxTL[1] & @CRLF & "Bottom Right of boxes [x] = " & $boxBR[0] & @CRLF & "Bottom Right of boxes [y] = " & $boxBR[1] & @CRLF & "Lowest point of case [x] = " & $bottom[0] & @CRLF & "Lowest point of case [y] = " & $bottom[1] & @CRLF & "Top Left of error banner area [x] = " & $errorTL[0] & @CRLF & "Top Left of error banner area [y] = " & $errorTL[1] & @CRLF & "Bottom Right of error banner area [x] = " & $errorBR[0] & @CRLF & "Bottom Right of error banner [y] = " & $errorBR[1] & @CRLF & "Center of next button [x] = " & $next[0] & @CRLF & "Center of next button [y] = " & $next[1] & @CRLF & "Bottom quarter of next button [x] = " & $nextButton[0] & @CRLF & "Bottom quarter of next button [y] = " & $nextButton[1] & @CRLF & "Location of 'Pick One' button [x] = " & $pickButton[0] & @CRLF & "Location of 'Pick One' button [y] = " & $pickButton[1] & @CRLF & "Color of Pop Now figure = 0x" & Hex($iColorBox, 6))

	FileClose($rFileHandle)
EndFunc

Func SavePreset()
	$sNewPreset = InputBox("Save Preset", "Enter the name of the preset.")

	; Create file in same folder as scripts
    $sFileName = "C:\Pop Now Bot Presets\" & $sNewPreset & ".txt"
	
	If FileExists($sFileName) Then
	    If (MsgBox($MB_YESNO, "File", "File already exists. Repalce old file?") = 6) Then
			$hFilehandle = FileOpen($sFileName, $FO_OVERWRITE)
		EndIf
	EndIf

	; Open file
	$hFilehandle = FileOpen($sFileName, $FO_CREATEPATH)

	$hFilehandle = FileOpen($sFileName, $FO_APPEND)
	
	; Prove it exists
	If Not FileExists($sFileName) Then
	    MsgBox($MB_SYSTEMMODAL, "File", "Error creating file")
	EndIf

    ; Write data to the file using the handle returned by FileOpen.
    FileWrite($hFilehandle, $boxTL[0] & @CRLF & $boxTL[1] & @CRLF & $boxBR[0] & @CRLF & $boxBR[1] & @CRLF & $bottom[0] & @CRLF & $bottom[1] & @CRLF & $errorTL[0] & @CRLF &  $errorTL[1] & @CRLF & $errorBR[0] & @CRLF & $errorBR[1] & @CRLF & $next[0] & @CRLF & $next[1] & @CRLF & $nextButton[0] & @CRLF & $nextButton[1] & @CRLF & $pickButton[0] & @CRLF & $pickButton[1] & @CRLF & $iColorBox)

	MsgBox($MB_SYSTEMMODAL, "File Content", FileRead($sFileName))

    ; Close the handle returned by FileOpen.
    FileClose($sFileName)

	MsgBox($MB_OK, "Save Preset", "Preset saved." & @CRLF & "File location: " & $sFileName)

EndFunc

Func Terminate()
	Exit
EndFunc ;== Terminate