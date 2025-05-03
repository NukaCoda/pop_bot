;version 14
#include <WinAPIGdi.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>

Global $g_bPaused = False, $iColorBox = 0xFF00FF, $iColorBottom = 0xFF00FF, $iColorNext = 0x000000, $mousePos = [0,0]
Global $boxTL = [0,0], $boxBR = [0,0], $next = [0,0], $bottom = [0,0], $pickButton = [0,0], $nextButton = [0,0], $error = [0,0]
Global $iTimeout = 10
Global $scale = _WinAPI_EnumDisplaySettings('', $ENUM_CURRENT_SETTINGS)[0] / @DesktopWidth
Global $casesPassed = 0, $resumeAfterFind = 0
Global $errorChecksum, $nextChecksum, $caseChecksum
Global $debugTime = 200

HotKeySet("{HOME}", "TogglePause")
HotKeySet("{ESC}", "Terminate")

$ready = MsgBox($MB_OKCANCEL, "Pop Now Bot", "This program will automatically click through pages of Pop Now figures and click the 'Pick One to Shake' button when one is available. There are some steps to ensure the program runs smoothly. Ready to set up the parameters for Pop Now Bot?")

If ($ready = 1) Then
	If (MsgBox($MB_YESNO, "Load Preset", "Do you want to load a preset?") = 7) Then
		Local $continue = 10
		;== Set TL bounds for boxes
		While ($continue = 10)
			If (MsgBox($MB_OKCANCEL, "Box Locations", "Position your cursor at the upper left bounds of where the Pop Now boxes are located on your screen. Position will be saved in " & $iTimeout & "seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIF
			$boxTL = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Box Locations", "Top left boxes position set to x = " & $boxTL[0] & ", y = " & $boxTL[1] & ". Continue?")
		WEnd
		;== If responce is 'Cancel', quit program
		If ($continue = 2) Then
			Terminate()
		EndIf

		;== Reset response
		$continue = 10
		;== Set BR bounds for boxes
		While ($continue = 10)
			If (MsgBox($MB_OKCANCEL, "Box Locations", "Position your cursor at the bottom right bounds of where the Pop Now boxes are located on your screen. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIf
			$boxBR = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Box Locations", "Bottom right boxes position set to x = " & $boxBR[0] & ", y = " & $boxBR[1] & ". Continue?")
		WEnd
		;== If responce is 'Cancel', quit program
		If ($continue = 2) Then
			Terminate()
		EndIf

		;== Reset response
		$continue = 10
		;== Set position of bottom of case
		While ($continue = 10)
			If (MsgBox($MB_OKCANCEL, "Case Locations", "Position your cursor on a unique color on the bottom of the case. Position will be saved in " & $iTimeout & " seconds or click cancel.",$iTimeout) = 2) Then
				Terminate()
			EndIf
			$bottom = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Case Locations", "Position set to x = " & $bottom[0] & ", y = " & $bottom[1] & ". Continue?")
		WEnd
		;== If responce is 'Cancel', quit program
		If ($continue = 2) Then
			Terminate()
		EndIf

		;== Reset response
		$continue = 10
		;== Set position for error message
		While ($continue = 10)
			If (MsgBox($MB_OKCANCEL, "Error Locations", "Position your cursor in the center of the banner above the 'Shake for Hints' text. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIf
			$error = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Error Locations", "Top left of error message position set to x = " & $error[0] & ", y = " & $error[1] & ". Continue?")
		WEnd
		;== If responce is 'Cancel', quit program
		If ($continue = 2) Then
			Terminate()
		EndIf

		;== Reset response
		$continue = 10
		;== Set position for 'next' checker
		While ($continue = 10)
			If (MsgBox($MB_OKCANCEL, "Next Case Checker", "Position your cursor in the center of the 'Next' button. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIf
			$next = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Next Case Checker", "Position set to x = " & $next[0] & ", y = " & $next[1] & ". Continue?")
		WEnd
		;== If responce is 'Cancel', quit program
		If ($continue = 2) Then
			Terminate()
		EndIf

		;== Reset response
		$continue = 10
		;== set position of 'next' button
		While ($continue = 10) 
			If (MsgBox($MB_OKCANCEL, "Button Positions", "Position your cursor in the bottom quarter of the 'Next Case' button. Position will be saved in " & $iTimeout & " seconds or click cancel.",$iTimeout) = 2) Then
				Terminate()
			EndIf
			$nextButton = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Button Positions", "Position of 'Next Case' button set to x = " & $nextButton[0] & ", y = " & $nextButton[1] & ". Continue?")
		WEnd
		;== If responce is 'Cancel', quit program
		If ($continue = 2) Then
			Terminate()
		EndIf

		;== Reset response
		$continue = 10
		;== Set position of 'pick one' button
		While ($continue = 10)
			If (MsgBox($MB_OKCANCEL, "Button Positions", "Position your cursor on the 'Pick One to Shake' button. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIf
			$pickButton = MouseGetPos()
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Button Positions", "Position of 'Pick One to Shake' button set to x = " & $pickButton[0] & ", y = " & $pickButton[1] & ". Continue?")
		WEnd
		;== If responce is 'Cancel', quit program
		If ($continue = 2) Then
			Terminate()
		EndIf

		;== Reset response
		$continue = 10
		;== Set color to search for available box
		While ($continue = 10)
			If (MsgBox($MB_OKCANCEL, "Color Selection", "Position your cursor on a UNIQUE color of a Pop Now figure. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
				Terminate()
			EndIf
			$mousePos = MouseGetPos()
			$iColorBox = PixelGetColor($mousePos[0], $mousePos[1])
			$continue = MsgBox($MB_CANCELTRYCONTINUE, "Color Selection", "Hex color code for Pop Now box color is 0x" & Hex($iColorBox, 6) & ". Continue?")
		WEnd
		;== If responce is 'Cancel', quit program
		If ($continue = 2) Then
			Terminate()
		EndIf
	Else
		LoadPreset()
	EndIf
Else
	Terminate()
EndIf ;== Finish setup *************************************************************************

Sleep(500)

$mouseTemp = MouseGetPos()
MouseMove(0, 0, 0)

;== Error message location
$errorChecksum = PixelChecksum($error[0] - 10, $error[1] - 10, $error[0] + 10, $error[1] + 10)
;== Next loaded location
$nextChecksum = PixelChecksum($next[0] - 10, $next[1] - 10, $next[0] + 10, $next[1] + 10)
;== Case in position location
$caseChecksum = PixelChecksum($bottom[0] - 1, $bottom[1] - 1, $bottom[0] + 1, $bottom[1] + 1)
;== Left of case color
$boxTLcolor = PixelGetColor($boxTL[0], $boxTL[1])

MouseMove($mouseTemp[0], $mouseTemp[1], 0)

Sleep(500)

;== Save Preset
If (MsgBox($MB_YESNO, "Save Preset", "Would you like to save these parameters as a preset?") = 6) Then
	SavePreset()
EndIf

MsgBox($MB_OK, "Pop Now Bot", "Set up complete! Press the 'Home' key to start and stop the program. Press the 'Escape' key to close or kill the program." & @CRLF & @CRLF & "Happy hunting!")

While 1
	ToolTip("Pop Now Bot" & @CRLF & "PAUSED" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
			Sleep(50)
WEnd
	
Func TogglePause()
	$g_bPaused = Not $g_bPaused
	If $resumeAfterFind = 1 Then
		$casesPassed = 0
	EndIf
	;== Program paused
	While Not $g_bPaused
		If $casesPassed = 1 Then
			ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " case passed - PAUSED" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
			Sleep(50)
		Else
			ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " cases passed - PAUSED" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
			Sleep(50)
		EndIf
	Wend ;== Program paused

	;== Program unpaused
	While $g_bPaused
		If $casesPassed = 1 Then
			ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " case passed" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
			Sleep(50)
		Else
			ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " cases passed" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
			Sleep(50)
		EndIf

		;== Next check
		;== Next button loaded
		If ($nextChecksum = PixelChecksum($next[0] - 10, $next[1] - 10, $next[0] + 10, $next[1] + 10)) Then
			;== Case check
			;== Case in position
			If ($caseChecksum = PixelChecksum($bottom[0] - 1, $bottom[1] - 1, $bottom[0] + 1, $bottom[1] + 1)) Then
				;== Boxes color check
				Sleep(100)
				$boxes = PixelSearch($boxTL[0], $boxTL[1], $boxBR[0], $boxBR[1], $iColorBox, 0)
				;== Box check
				;== Box available
				If IsArray($boxes) Then
					;== Mouse clicks pick button
					MouseMove($pickButton[0], $pickButton[1], 0)
					MouseClick($MOUSE_CLICK_LEFT)
					$errorPresent = 0
					$errorTimer = 0

					;== False positive check
					While ($errorTimer <= 100)
						ToolTip($errorTimer & "% Lie Detecting")
						Sleep(20)
						ToolTip($errorTimer & "% Lie Detecting")
						Sleep(20)
						ToolTip($errorTimer & "% Lie Detecting")
						;== If no error message, errorPresent stays 0
						If ($errorChecksum = PixelChecksum($error[0] - 10, $error[1] - 10, $error[0] + 10, $error[1] + 10)) Then
							$errorTimer = $errorTimer + 2
						;== If error message, errorPresent counts up
						Else
							$errorPresent = $errorPresent + 1
							$errorTimer = $errorTimer + 2
						EndIf
						$g_bPaused = 0
						Sleep(10)
					Wend
					;== Error is present, false positive
					If ($errorPresent > 0) Then
						$g_bPaused = 1
						$resumeCounter = 400
						;== Program continues
						While $resumeCounter > 0
							ToolTip("LIE DETECTED" & @CRLF & "Resuming search in " & Floor($resumeCounter / 100) + 1)
							$resumeCounter = $resumeCounter - 1
							Sleep(1)
						Wend
						$casesPassed = $casesPassed + 1
						MouseMove($nextButton[0], $nextButton[1], 0)
						MouseClick($MOUSE_CLICK_LEFT)
						Sleep(2000)
					;== Error not present, box successfully found
					Else
						$resumeAfterFind = 1
						$g_bPaused = 0
						;== Program pauses until user resumes or quits
						While Not $g_bPaused
							If $casesPassed = 0 Then
								ToolTip("Pop Now Bot" & @CRLF & "Box found!" & @CRLF & "It only took " & $casesPassed + 1 & " case!")
							Else
								ToolTip("Pop Now Bot" & @CRLF & "Box found!" & @CRLF & "It only took " & $casesPassed + 1 & " cases!")
							EndIf
							Sleep(50)
						Wend
					EndIf
				;== Box unavailable
				Else
					MouseMove($nextButton[0], $nextButton[1], 0)
					$ShouldClick = 1
					;== Mouse clicks once
					While $ShouldClick = 1
						MouseClick($MOUSE_CLICK_LEFT)
						$ShouldClick = 0
					WEnd
					;== Waits for case to move
					While $ShouldClick = 0
						;== Case has not moved yet
						If $boxTLcolor = PixelGetColor($boxTL[0], $boxTL[1]) Then
							;== Do nothing
						;== Case moves
						Else
							;== Reset ShouldClick
							$ShouldClick = 1
							$casesPassed = $casesPassed + 1
						EndIF
					WEnd
					$boxes = 0
					$case = 0
					$loaded = 0
				EndIf ;== Box check
				$boxes = 0
				$case = 0
				$loaded = 0
			;== Case not in position
			EndIf ;== Case check
			$boxes = 0
			$case = 0
			$loaded = 0
		;== Next button not loaded
		EndIf ;== Next check
		$boxes = 0
		$case = 0
		$loaded = 0
	WEnd ;== Program unpaused
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
	$error[0]		= FileReadLine($rFileHandle, 7)
	$error[1]		= FileReadLine($rFileHandle, 8)
	$next[0]		= FileReadLine($rFileHandle, 9)
	$next[1]		= FileReadLine($rFileHandle, 10)
	$nextButton[0]	= FileReadLine($rFileHandle, 11)
	$nextButton[1]	= FileReadLine($rFileHandle, 12)
	$pickButton[0]	= FileReadLine($rFileHandle, 13)
	$pickButton[1]	= FileReadLine($rFileHandle, 14)
	$iColorBox		= FileReadLine($rFileHandle, 15)

	MsgBox($MB_SYSTEMMODAL, "Loaded Content", "Top Left of boxes [x] = " & $boxTL[0] & @CRLF & "Top Left of boxes [y] = " & $boxTL[1] & @CRLF & "Bottom Right of boxes [x] = " & $boxBR[0] & @CRLF & "Bottom Right of boxes [y] = " & $boxBR[1] & @CRLF & "Lowest point of case [x] = " & $bottom[0] & @CRLF & "Lowest point of case [y] = " & $bottom[1] & @CRLF & "Center of error banner area [x] = " & $error[0] & @CRLF & "Center of error banner area [y] = " & $error[1] & @CRLF & "Center of next button [x] = " & $next[0] & @CRLF & "Center of next button [y] = " & $next[1] & @CRLF & "Bottom quarter of next button [x] = " & $nextButton[0] & @CRLF & "Bottom quarter of next button [y] = " & $nextButton[1] & @CRLF & "Location of 'Pick One' button [x] = " & $pickButton[0] & @CRLF & "Location of 'Pick One' button [y] = " & $pickButton[1] & @CRLF & "Color of Pop Now figure = 0x" & Hex($iColorBox, 6))

	FileClose($rFileHandle)
EndFunc ;== LoadPreset

Func SavePreset()
	$sNewPreset = InputBox("Save Preset", "Enter the name of the preset.")

	; Create file in same folder as scripts
    $sFileName = "C:\Pop Now Bot Presets\" & $sNewPreset & ".txt"
	
	If FileExists($sFileName) Then
	    If (MsgBox($MB_YESNO, "File", "File already exists. Repalce old file?") = 6) Then
			$hFilehandle = FileOpen($sFileName, $FO_OVERWRITE)
		EndIf
	EndIf

	$sFilePath = "C:\Pop Now Bot Presets"

	DirCreate($sFilePath)

	; Open file
	FileOpen($sFileName, $FO_CREATEPATH)

	FileClose($sFileName)
	$hFilehandle = FileOpen($sFileName, $FO_APPEND)
	
	; Prove it exists
	If Not FileExists($sFileName) Then
	    MsgBox($MB_SYSTEMMODAL, "File", "Error creating file." & @CRLF & @CRLF & "Application closing")
		Terminate()
	EndIf

    ; Write data to the file using the handle returned by FileOpen.
    FileWrite($hFilehandle, $boxTL[0] & @CRLF & $boxTL[1] & @CRLF & $boxBR[0] & @CRLF & $boxBR[1] & @CRLF & $bottom[0] & @CRLF & $bottom[1] & @CRLF & $error[0] & @CRLF &  $error[1] & @CRLF & $next[0] & @CRLF & $next[1] & @CRLF & $nextButton[0] & @CRLF & $nextButton[1] & @CRLF & $pickButton[0] & @CRLF & $pickButton[1] & @CRLF & $iColorBox)

	MsgBox($MB_SYSTEMMODAL, "File Content", FileRead($sFileName))

    ; Close the handle returned by FileOpen.
    FileClose($sFileName)

	MsgBox($MB_OK, "Save Preset", "Preset saved." & @CRLF & "File location: " & $sFileName)

EndFunc ;== SavePreset

Func Terminate()
	Exit
EndFunc ;== Terminate