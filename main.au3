;version 25.1
#include <WinAPIGdi.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>
#include <ScreenCapture.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Color.au3>
#Include <Misc.au3>

Global $g_bPaused = False, $casesPassed = 0, $resumeAfterFind = 0, $nextParameter = 500, $successVolume = 10, $soundPlayed = 0
Global $refreshTimer = 1, $timeoutTimer = 400, $timeoutReset = 400
Global $iColorBox = 0xFF00FF, $iColorBottom = 0xFF00FF, $iColorNext = 0x000000, $iColorBlack = 0x000000, $iColorWhite = 0xFFFFFF, $iColorMagenta = 0xFF00FF, $iColorCase, $iColorCaseLeft, $iColorPick
Global $box1 = [0,0], $box2 = [0,0], $next = [0,0], $bottom = [0,0], $pickButton = [0,0], $nextButton = [0,0], $error = [0,0], $caseLeft = [0, 0], $mousePos = [0,0]
Global $iTimeout = 10, $bottomOffset = 10, $nextOffset = 15, $errorOffset = 10, $pickOffset = 5
Global $errorChecksum, $nextChecksum, $caseChecksum
Global $iX1, $iY1, $iX2, $iY2, $aPos, $sMsg, $sBMP_Path, $hRectangle_GUI

Global $debugTime = 200

HotKeySet("{HOME}", "TogglePause")
HotKeySet("{ESC}", "Terminate")

Initiate()

Func Initiate()
	Local $autoRefresh = 1
	$ready = MsgBox($MB_OKCANCEL, "Pop Bot", "This program will automatically click through pages of Pop Now figures and click the 'Pick One to Shake' button when one is available. There are some steps to ensure the program runs smoothly. Ready to set up the parameters for Pop Bot?")

	If ($ready = 1) Then
		$presetResponse = MsgBox($MB_YESNO, "Load Preset", "Do you want to load a preset?")
		If $presetResponse = 6 Then
			LoadPreset()
			$autoRefresh = PreDrop()
		ElseIf $presetResponse = 7 Then
			Setup()
		EndIf
	Else
		Terminate()
	EndIf
	If $autoRefresh = 1 Then
		SavePreset()

		MsgBox($MB_OK, "Pop Bot", "Set up complete! Press the 'Home' key to start and stop the program. Press the 'Escape' key to close or kill the program." & @CRLF & @CRLF & "Happy hunting!")

		While 1
			ToolTip("Pop Bot" & @CRLF & @CRLF & @CRLF & "Home to start" & @CRLF & "Escape to quit")
			Sleep(50)
		WEnd
	ElseIf $autoRefresh = 0 Then
		$g_bPaused = 0
		TogglePause()
	EndIf
EndFunc ;== Initiate

Func ShowParameters()
	$mouseTemp = MouseGetPos()
	Local $counterPreset = 0

	Do
		MouseMove($box1[0], $box1[1])
		ToolTip("Box position 1")
		$counterPreset = $counterPreset + 10
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($box2[0], $box2[1])
		ToolTip("Box position 2")
		$counterPreset = $counterPreset + 10
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($bottom[0], $bottom[1])
		ToolTip("Case position")
		$counterPreset = $counterPreset + 10
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($error[0], $error[1])
		ToolTip("Error position")
		$counterPreset = $counterPreset + 20
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($next[0], $next[1])
		ToolTip("Next button position")
		$counterPreset = $counterPreset + 10
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($nextButton[0], $nextButton[1])
		ToolTip("Next click position")
		$counterPreset = $counterPreset + 10
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($pickButton[0], $pickButton[1])
		ToolTip("Pick button position")
		$iColorPick = PixelGetColor($pickButton[0], $pickButton[1])
		;== Error message location
		$errorChecksum = PixelChecksum($error[0] - $errorOffset, $error[1] - $errorOffset, $error[0] + $errorOffset, $error[1] + $errorOffset)
		;== Next loaded location
		$nextChecksum = PixelChecksum($next[0] - $nextOffset, $next[1] - $nextOffset, $next[0] + $nextOffset, $next[1] + $nextOffset)
		;== Case in position location
		$caseChecksum = PixelChecksum($bottom[0] - $bottomOffset, $bottom[1] - $bottomOffset, $bottom[0] + 	$bottomOffset, $bottom[1] + $bottomOffset)
		$iColorCase = PixelGetColor($bottom[0], $bottom[1])
		;== Left of case color
		$iColorCaseLeft = PixelGetColor($caseLeft[0], $caseLeft[1])
		$counterPreset = $counterPreset + 10
		Sleep(1)
	Until $counterPreset >= $nextParameter
	ToolTip("")
	MouseMove($mouseTemp[0], $mouseTemp[1])

	GUICreate("", 200, 200) ; will create a dialog box that when displayed is centered
    GUISetBkColor($iColorBox)
	$textRGB = _ColorGetRGB($iColorBox)

	$brightText = GUICtrlCreateLabel("Selected box color" & @CRLF & @CRLF & "Close this window to proceed", 10, 10)
	
	If $textRGB[0] < 128 or $textRGB[1] < 128 or $textRGB[2] < 128 Then
	GUICtrlSetColor($brightText, 0xFFFFFF)
	EndIf

    GUISetState(@SW_SHOW)

    ; Loop until the user exits.
    While 1
            Switch GUIGetMsg()
                    Case $GUI_EVENT_CLOSE
                            ExitLoop
            EndSwitch
    WEnd

	GUIDelete()
EndFunc ;== ShowParameters

Func PreDrop()
	Local $nextPresent
	Local $confirm = MsgBox($MB_YESNO, "Pop Bot", "Do you want the bot to automatically refresh the page in anticipation for a Pop Now restock?")

	If $confirm = 6 Then
		Local $title = WinGetTitle("POP NOW", "")
		
		$refreshTimer = InputBox("Pop Bot", "Enter number of minutes to wait in between each refresh")

		Local $refresh = $refreshTimer * 120000

		Do
			If WinExists($title) and $refresh < 1 Then
				WinActivate($title)
				WinWaitActive($title)
				Send("{F5}")
				$refresh = $refreshTimer * 120000
			EndIf

			If $refresh < 0 Then
				$title = WinGetTitle("POP MART", "")
			ElseIf $refresh > (($refreshTimer * 120000) - 1) Then
				$title = WinGetTitle("POP NOW", "")
			EndIf

			If Floor($refresh * 50 / 100000) = 1 Then
				ToolTip("Pop Bot" & @CRLF & Floor($refresh * 50 / 100000) & " second until next refresh")
			Else
				ToolTip("Pop Bot" & @CRLF & Floor($refresh * 50 / 100000) & " seconds until next refresh")
			EndIf
			Sleep(50)

			$nextPresent = PixelSearch($next[0] - 15, $next[1] - 15, $next[0] + 15, $next[1] + 15, $iColorBlack, 190)

			$refresh = $refresh - 125
		Until IsArray($nextPresent)
		Sleep(5000)
		Return 0
	ElseIf $confirm = 7 Then
		Return 1
	EndIf
EndFunc

Func TogglePause()
	$g_bPaused = Not $g_bPaused
	If $resumeAfterFind = 1 Then
		SoundSetWaveVolume($successVolume)
		SoundPlay("", 0)
		$soundPlayed = 0
		$g_bPaused = 0
		$casesPassed = 0
		$resumeAfterFind = 0
	EndIf
	;== Program paused
	While $g_bPaused = 0
		If $casesPassed = 1 Then
			ToolTip("Pop Bot" & @CRLF & $casesPassed & " case passed - PAUSED" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
			Sleep(50)
		Else
			ToolTip("Pop Bot" & @CRLF & $casesPassed & " cases passed - PAUSED" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
			Sleep(50)
		EndIf
	WEnd ;== Program paused

	;== Program unpaused
	While $g_bPaused = 1
		If $casesPassed = 1 Then
			ToolTip("Pop Bot" & @CRLF & $casesPassed & " case passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
			Sleep(50)
		Else
			ToolTip("Pop Bot" & @CRLF & $casesPassed & " cases passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
			Sleep(50)
		EndIf

		;== Next check
		;== Next button loaded
		;ToolTip("Next check")
		;Sleep(100)

		While PixelGetColor($pickButton[0], $pickButton[1]) = $iColorWhite and $timeoutTimer > 0
			ToolTip("Pop Bot" & @CRLF & Floor($timeoutTimer / 20) & " until timeout refresh")
			$timeoutTimer = $timeoutTimer - 1
			Sleep(50)
		WEnd

		If $timeoutTimer < 1 Then
			Send("{F5}")
			$timeoutTimer = $timeoutReset
		Else
			$timeoutTimer = $timeoutReset
		EndIf

		If ($nextChecksum = PixelChecksum($next[0] - 15, $next[1] - 15, $next[0] + 15, $next[1] + 15)) Then
			;== Case check
			;== Case in position
			;ToolTip("Case check")
			;Sleep(100)
			If ($caseChecksum = PixelChecksum($bottom[0] - $bottomOffset, $bottom[1] - $bottomOffset, $bottom[0] + $bottomOffset, $bottom[1] + $bottomOffset)) Then
				;== Boxes color check
				;Sleep(100)
				$boxes = PixelSearch($box1[0], $box1[1], $box2[0], $box2[1], $iColorBox, 0)
				;== Box check
				;== Box available
				;ToolTip("Box check")
				;Sleep(100)
				If IsArray($boxes) Then
					;== Mouse clicks pick button
					MouseMove($pickButton[0], $pickButton[1], 0)
					MouseClick($MOUSE_CLICK_LEFT)

					If $casesPassed = 1 Then
						ToolTip("Pop Bot" & @CRLF & $casesPassed & " case passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					Else
						ToolTip("Pop Bot" & @CRLF & $casesPassed & " cases passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					EndIf

					$errorPresent = 0
					$errorTimer = 0
					;$loadingStage = 0
					Local $pickLoad = 0

					;== False positive check
					While ($errorTimer <= 10000)
						While ($errorTimer <= 20)
							ToolTip("Lie Detecting - " & Floor($errorTimer / 100) & "% complete")
							;== Color of pick button when it is present with the cursor over it
							$pickLoad = PixelGetColor($pickButton[0] + 5, $pickButton[1] - 5)
							$errorTimer = $errorTimer + 1
							Sleep(50)
						WEnd
						While ($errorTimer <= 9999)
							If (PixelGetColor($pickButton[0] + 5, $pickButton[1] - 5) = $iColorPick) or (PixelGetColor($pickButton[0] + 5, $pickButton[1] - 5) = $pickLoad) and $errorPresent = 0 Then
								ToolTip("Lie Detecting - " & Floor($errorTimer / 100) & "% complete")
								;== If no error message, errorPresent stays 0
								If ($errorChecksum = PixelChecksum($error[0] - 10, $error[1] - 10, $error[0] + 10, $error[1] + 10)) Then
									$errorTimer = $errorTimer + 1
								;== If error message, errorPresent counts up
								Else
									$errorPresent = $errorPresent + 1
									$errorTimer = $errorTimer + 1
								EndIf
								Sleep(50)
							ElseIf PixelGetColor($pickButton[0] + 5, $pickButton[1] - 5) = $iColorWhite and $errorPresent = 0 Then
								$iColorPick = $iColorMagenta
								$pickLoad = $iColorMagenta
								ToolTip("Lie Detecting - " & Floor($errorTimer / 100) & "% complete")
								;== If no error message, errorPresent stays 0
								If ($errorChecksum = PixelChecksum($error[0] - 10, $error[1] - 10, $error[0] + 10, $error[1] + 10)) Then
									$errorTimer = $errorTimer + 10
								;== If error message, errorPresent counts up
								Else
									$errorPresent = $errorPresent + 1
									$errorTimer = $errorTimer + 10
								EndIf
								Sleep(50)
							Else
								$iColorPick = $iColorMagenta
								$pickLoad = $iColorMagenta
								ToolTip("Lie Detecting - " & Floor($errorTimer / 100) & "% complete")
								;== If no error message, errorPresent stays 0
								If ($errorChecksum = PixelChecksum($error[0] - 10, $error[1] - 10, $error[0] + 10, $error[1] + 10)) Then
									$errorTimer = $errorTimer + 500
								;== If error message, errorPresent counts up
								Else
									$errorPresent = $errorPresent + 1
									$errorTimer = $errorTimer + 500
								EndIf
								Sleep(50)
							EndIf
							$errorTimer = $errorTimer + 20
						WEnd
						$errorTimer = $errorTimer + 20
					WEnd
					;== Error is present, false positive
					If ($errorPresent > 0) Then
						$g_bPaused = 1
						$resumeCounter = 400
						;== Program continues
						While $resumeCounter > 0
							ToolTip("LIE DETECTED" & @CRLF & "Resuming search in " & Floor($resumeCounter / 100) + 1)
							$resumeCounter = $resumeCounter - 5
							Sleep(50)
						WEnd
						$casesPassed = $casesPassed + 1
						MouseMove($nextButton[0], $nextButton[1], 0)
						MouseClick($MOUSE_CLICK_LEFT)
						If $casesPassed = 1 Then
						ToolTip("Pop Bot" & @CRLF & $casesPassed & " case passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					Else
						ToolTip("Pop Bot" & @CRLF & $casesPassed & " cases passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					EndIf
						Sleep(2000)
					;== Error not present, box successfully found
					Else
						$resumeAfterFind = 1
						$g_bPaused = 0
						;== Program pauses until user resumes or quits
						While $g_bPaused = 0
							If $casesPassed = 0 Then
								ToolTip("Pop Bot" & @CRLF & "Box found!" & @CRLF & "It only took " & $casesPassed + 1 & " case!")
							Else
								ToolTip("Pop Bot" & @CRLF & "Box found!" & @CRLF & "It only took " & $casesPassed + 1 & " cases!")
							EndIf
							If $soundPlayed = 0 Then
                                $soundPlayed = 1
                                SoundSetWaveVolume($successVolume)
							    SoundPlay("Success Sound\Champions.wav", 0)
							EndIf
                            Sleep(50)
						WEnd
					EndIf
				;== Box unavailable
				Else
					MouseMove($nextButton[0], $nextButton[1], 0)

					If $casesPassed = 1 Then
						ToolTip("Pop Bot" & @CRLF & $casesPassed & " case passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					Else
						ToolTip("Pop Bot" & @CRLF & $casesPassed & " cases passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					EndIf

					$ShouldClick = 1
					;== Mouse clicks once
					While $ShouldClick = 1
						MouseClick($MOUSE_CLICK_LEFT)
						$ShouldClick = 0
					WEnd
					;== Waits for case to move
					While $ShouldClick = 0
						;== Case has not moved yet
						If $iColorCaseLeft = PixelGetColor($box1[0], $box1[1]) Then
							;== Do nothing
						;== Case moves
						Else
							;== Reset ShouldClick
							$ShouldClick = 1
							$casesPassed = $casesPassed + 1
						EndIF
					WEnd
					$boxes = 0
				EndIf ;== Box check
				$boxes = 0
			;== Case not in position
			EndIf ;== Case check
			$boxes = 0
		;== Next button not loaded
		EndIf ;== Next check
		$boxes = 0
	WEnd ;== Program unpaused
	ToolTip("")
EndFunc ;== TogglePause

Func Setup()
	Local $continue = 10
	;== Set bounds for boxes
	; Create GUI prompting user to mark an area or cancel
	While $continue = 10
		$hMain_GUI = GUICreate("Boxes", 240, 50)

		$hRect_Button   = GUICtrlCreateButton("Mark Area",  10, 10, 80, 30)
		$hCancel_Button = GUICtrlCreateButton("Cancel", 150, 10, 80, 30)

		GUISetState()

		While 1

		    Switch GUIGetMsg()
		        Case $GUI_EVENT_CLOSE, $hCancel_Button
		            FileDelete(@TempDir & "\Rect.bmp")
					Terminate()
		        Case $hRect_Button
		            GUIDelete($hMain_GUI)
		            Mark_Rect()
		    ; Capture selected area
		            $sBMP_Path = @TempDir & "\Rect.bmp"
		            _ScreenCapture_Capture($sBMP_Path, $iX1, $iY1, $iX2, $iY2, False)
		    ; Display image
		            $hBitmap_GUI = GUICreate("Selected Rectangle", $iX2 - $iX1 + 1, $iY2 - $iY1 + 45, -1, -1)
					GUICtrlCreatePic($sBMP_Path, 0, 0, $iX2 - $iX1 + 1, $iY2 - $iY1 + 1)
		            $cancelButton = GUICtrlCreateButton("Cancel", 5, $iY2 - $iY1 + 10, (($iX2 - $iX1) / 3) - 10, 25)
					$retryButton = GUICtrlCreateButton("Try Again", (($iX2 - $iX1) / 3) + 5, $iY2 - $iY1 + 10, (($iX2 - $iX1) / 3) - 10, 25)
					$contButton = GUICtrlCreateButton("Continue", ((($iX2 - $iX1) / 3) * 2) + 5, $iY2 - $iY1 + 10, (($iX2 - $iX1) / 3) - 10, 25)
					GUISetState()
					WinActivate("Selected Rectangle")
					While 1
							Switch GUIGetMsg()
								Case $GUI_EVENT_CLOSE, $cancelButton
									Terminate()
								Case $contButton
									$continue = 11
									GUIDelete($hBitmap_GUI)
									ExitLoop
								Case $retryButton
									$continue = 10
									GUIDelete($hBitmap_GUI)
									ExitLoop
							EndSwitch
					WEnd
					ExitLoop
		    EndSwitch
		WEnd
		$box1[0] = $iX1
		$box1[1] = $iY1
		$box2[0] = $iX2
		$box2[1] = $iY2
	Wend

	If $box1[0] < $box2[0] Then
		$caseLeft[0] = $box1[0]
		$caseLeft[1] = $box1[1]
	Else
		$caseLeft[0] = $box2[0]
		$caseLeft[1] = $box2[1]
	EndIf

	;== Reset response
	$continue = 10
	;== Set position of bottom of case
	While ($continue = 10)
		If (MsgBox($MB_OKCANCEL, "Case Location", "Position your cursor on a unique color on the bottom of the case. Position will be saved in " & $iTimeout & " seconds or click cancel.", $iTimeout) = 2) Then
			Terminate()
		EndIf
		$bottom = MouseGetPos()

		$sCasePic_Path = @TempDir & "\Case.bmp"
		_ScreenCapture_Capture($sCasePic_Path, $bottom[0] - $bottomOffset, $bottom[1] - $bottomOffset, $bottom[0] + $bottomOffset, $bottom[1] + $bottomOffset, False)

		GUICreate("", $bottomOffset * 20, $bottomOffset * 20 + 110) ; will create a dialog box that when displayed is centered
		$colorCase = "0x" & Hex(PixelGetColor($bottom[0], $bottom[1]), 6)

		GUICtrlCreatePic($sCasePic_Path, $bottomOffset * 9, $bottomOffset * 9, $bottomOffset * 2, $bottomOffset * 2)

		; Display image

		$cancelButton = GUICtrlCreateButton("Cancel", 10, $bottomOffset * 21, ($bottomOffset * 20) - 20, 25)

		$retryButton = GUICtrlCreateButton("Try Again", 10, $bottomOffset * 21 + 30, ($bottomOffset * 20) - 20, 25)

		$contButton = GUICtrlCreateButton("Continue", 10, $bottomOffset * 21 + 60, ($bottomOffset * 20) - 20, 25)

		GUISetState()
		WinActivate("Selected Case Position")

		While 1
			Switch GUIGetMsg()
				Case $GUI_EVENT_CLOSE, $cancelButton
					Terminate()
				Case $contButton
					$continue = 11
					GUIDelete()
					ExitLoop
				Case $retryButton
					$continue = 10
					GUIDelete()
					ExitLoop
			EndSwitch
		WEnd
	WEnd
	;== If response is 'Cancel', quit program
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
	;== If response is 'Cancel', quit program
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
	;== If response is 'Cancel', quit program
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
	;== If response is 'Cancel', quit program
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
	;== If response is 'Cancel', quit program
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

		GUICreate("", 200, 200) ; will create a dialog box that when displayed is centered
		$colorCase = "0x" & Hex(PixelGetColor($mousePos[0], $mousePos[1]), 6)
    	GUISetBkColor($colorCase)
		$textRGB = _ColorGetRGB($colorCase)

		$brightText = GUICtrlCreateLabel("Selected box color" & @CRLF & @CRLF & "Close this window to proceed", 10, 10)
		
		If $textRGB[0] < 128 or $textRGB[1] < 128 or $textRGB[2] < 128 Then
		GUICtrlSetColor($brightText, 0xFFFFFF)
		EndIf

    	GUISetState(@SW_SHOW)

    	; Loop until the user exits.
    	While 1
    	        Switch GUIGetMsg()
    	                Case $GUI_EVENT_CLOSE
    	                        ExitLoop
    	        EndSwitch
    	WEnd

		GUIDelete()
		
		$iColorBox = PixelGetColor($mousePos[0], $mousePos[1])
		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Color Selection", "Hex color code for Pop Now box color is 0x" & Hex($iColorBox, 6) & ". Continue?")
	WEnd
	;== If response is 'Cancel', quit program
	If ($continue = 2) Then
		Terminate()
	EndIf
EndFunc ;== Setup

Func LoadPreset()
	$refreshGUI = 1
	Do
		GUICreate("Pop Bot Presets", 300, 360)
		Local $idPresetList = GUICtrlCreateList("", 25, 25, 250, 250)
		Local $idButton_Confirm = GUICtrlCreateButton("Load preset", 25, 275, 250, 25)
		Local $idButton_Delete = GUICtrlCreateButton("Delete Preset", 25, 310, 250, 25)

		Local $hSearch = FileFindFirstFile("Presets\*.txt")
		
		If $hSearch = -1 Then
    	    MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
    	    Return False
    	EndIf
		
		Local $sFileName = ""
		
		While 1
    	    $sFileName = FileFindNextFile($hSearch)
    	    ; If there is no more file matching the search.
    	    If @error Then ExitLoop
			
    	    ; Add file name to the list
    	    GUICtrlSetData($idPresetList, $sFileName)
    	WEnd
		
		FileClose($hSearch)
		
		GUICtrlSetState(-1, $GUI_FOCUS)
		
		GUISetState(@SW_SHOW)
		
		; Loop until the user exits.
    	While 1
    	    Switch GUIGetMsg()
    	        Case $GUI_EVENT_CLOSE
					Terminate()
    	            ExitLoop
    	        Case $idButton_Confirm
    	            MsgBox($MB_SYSTEMMODAL, "", "Selected " & GUICtrlRead($idPresetList) & " preset")
					$rFileName = "Presets\" & GUICtrlRead($idPresetList)
					$refreshGUI = 0
					ExitLoop
				Case $idButton_Delete
					If MsgBox($MB_OKCANCEL, "Delete Preset", "Are you sure you want to delete " & GUICtrlRead	($idPresetList) & " preset?") = 1 Then
						FileDelete("Presets\" & GUICtrlRead($idPresetList))
					EndIf
					ExitLoop
    	    EndSwitch
    	WEnd
		GuiDelete()
	Until $refreshGUI = 0

	$rFileHandle = FileOpen($rFileName, $FO_READ)

	If $rFileHandle = -1 Then
		MsgBox($MB_OK, "ERROR", "An error occurred when reading the file.")
		Terminate()
    EndIf

	$box1[0]		= FileReadLine($rFileHandle, 1)
	$box1[1]		= FileReadLine($rFileHandle, 2)
	$box2[0]		= FileReadLine($rFileHandle, 3)
	$box2[1]		= FileReadLine($rFileHandle, 4)
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

	FileClose($rFileHandle)

	ShowParameters()
EndFunc ;== LoadPreset

Func SavePreset()
	;== Ask
	If (MsgBox($MB_YESNO, "Save Preset", "Would you like to save these parameters as a preset?") = 7) Then
		Return
	EndIf

	$sNewPreset = InputBox("Save Preset", "Enter the name of the preset.")

	; Create file
    $sFileName = @WorkingDir & "\Presets\" & $sNewPreset & ".txt"
	
	If FileExists($sFileName) Then
	    If (MsgBox($MB_YESNO, "File", "File already exists. Replace old file?") = 6) Then
			$hFilehandle = FileOpen($sFileName, $FO_OVERWRITE)
		EndIf
	EndIf

	$sFilePath = "Presets\"

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
    FileWrite($hFilehandle, $box1[0] & @CRLF & $box1[1] & @CRLF & $box2[0] & @CRLF & $box2[1] & @CRLF & $bottom[0] & @CRLF & $bottom[1] & @CRLF & $error[0] & @CRLF &  $error[1] & @CRLF & $next[0] & @CRLF & $next[1] & @CRLF & $nextButton[0] & @CRLF & $nextButton[1] & @CRLF & $pickButton[0] & @CRLF & $pickButton[1] & @CRLF & $iColorBox)

	;MsgBox($MB_SYSTEMMODAL, "File Content", FileRead($sFileName))

    ; Close the handle returned by FileOpen.
    FileClose($sFileName)

	MsgBox($MB_OK, "Save Preset", "Preset saved." & @CRLF & "File location: " & FileGetLongName($sFileName))
EndFunc ;== SavePreset

Func Mark_Rect()
    Local $aMouse_Pos, $aMask, $aM_Mask, $iTemp
    Local $UserDLL = DllOpen("user32.dll")

	; Wait until mouse button pressed
    While Not _IsPressed("01", $UserDLL)
		ToolTip("Pop Bot" & @CRLF & @CRLF & "Click and Drag to select area")
        Sleep(50)
    WEnd

	; Get first mouse position
    $aMouse_Pos = MouseGetPos()
    $iX1 = $aMouse_Pos[0]
    $iY1 = $aMouse_Pos[1]

    $hRectangle_GUI = GUICreate("", @DesktopWidth, @DesktopHeight, 0, 0, $WS_POPUP, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
    GUISetBkColor(0x000000)

	; Draw rectangle while mouse button pressed
    While _IsPressed("01", $UserDLL)
		ToolTip("Pop Bot" & @CRLF & @CRLF & "Selecting area")
        $aMouse_Pos = MouseGetPos()

        $aM_Mask = DllCall("gdi32.dll", "long", "CreateRectRgn", "long", 0, "long", 0, "long", 0, "long", 0)
	; Bottom of rectangle
        $aMask = DllCall("gdi32.dll", "long", "CreateRectRgn", "long", $iX1, "long", $aMouse_Pos[1], "long", $aMouse_Pos[0], "long", $aMouse_Pos[1] + 1)
        DllCall("gdi32.dll", "long", "CombineRgn", "long", $aM_Mask[0], "long", $aMask[0], "long", $aM_Mask[0], "int", 2)
	; Left of rectangle
        $aMask = DllCall("gdi32.dll", "long", "CreateRectRgn", "long", $iX1, "long", $iY1, "long", $iX1 + 1, "long", $aMouse_Pos[1])
        DllCall("gdi32.dll", "long", "CombineRgn", "long", $aM_Mask[0], "long", $aMask[0], "long", $aM_Mask[0], "int", 2)
	; Top of rectangle
        $aMask = DllCall("gdi32.dll", "long", "CreateRectRgn", "long", $iX1 + 1, "long", $iY1 + 1, "long", $aMouse_Pos[0], "long", $iY1)
        DllCall("gdi32.dll", "long", "CombineRgn", "long", $aM_Mask[0], "long", $aMask[0], "long", $aM_Mask[0], "int", 2)
	; Right of rectangle
        $aMask = DllCall("gdi32.dll", "long", "CreateRectRgn", "long", $aMouse_Pos[0], "long", $iY1, "long", $aMouse_Pos[0] + 1, "long", $aMouse_Pos[1])
        DllCall("gdi32.dll", "long", "CombineRgn", "long", $aM_Mask[0], "long", $aMask[0], "long", $aM_Mask[0], "int", 2)
        DllCall("user32.dll", "long", "SetWindowRgn", "hwnd", $hRectangle_GUI, "long", $aM_Mask[0], "int", 1)

		If BitAND(WinGetState(''), 2) <> 0 Then GUISetState()

        Sleep(50)

    WEnd
	ToolTip("")
	; Get second mouse position
    $iX2 = $aMouse_Pos[0]
    $iY2 = $aMouse_Pos[1]

	; Set in correct order if required
    If $iX2 < $iX1 Then
        $iTemp = $iX1
        $iX1 = $iX2
        $iX2 = $iTemp
    EndIf
    If $iY2 < $iY1 Then
        $iTemp = $iY1
        $iY1 = $iY2
        $iY2 = $iTemp
    EndIf

    GUIDelete($hRectangle_GUI)
    DllClose($UserDLL)
EndFunc ;== Mark_Rect

Func Terminate()
	Exit
EndFunc ;== Terminate