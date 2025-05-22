;version 30.2
#include <Array.au3>
#include <WinAPIGdi.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>
#include <ScreenCapture.au3>
#include <GUIConstantsEx.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <Color.au3>
#Include <Misc.au3>
#include <GDIPlus.au3>
#include "Libraries\UWPOCR.au3"

;True/False
Global $g_bPaused = False, $freshStart = true, $resumeAfterFind = 0, $soundPlayed = 0
;Timers & Counters
Global $iTimeout = 10, $timeoutTimer = 200, $timeoutReset = 200, $serialCounter = 0, $SNappearanceCount = 1, $casesPassed = 0, $uniquePassed = 0, $dupePassed = 0
;Colors
Global $iColorBox = 0xFF00FF, $iColorBottom = 0xFF00FF, $iColorNext = 0x000000, $iColorBlack = 0x000000, $iColorWhite = 0xFFFFFF, $iColorMagenta = 0xFF00FF, $iColorRed = 0xFF0000, $iColorGreen = 0x00FF00, $iColorCase, $iColorCaseLeft, $iColorPick
;Coordinates/Arrays
Global $box1 = [0,0], $box2 = [0,0], $next = [0,0], $bottom = [0,0], $pickButton = [0,0], $nextButton = [0,0], $error = [0,0], $caseLeft = [0, 0], $mousePos = [0,0], $serial1 = [0,0], $serial2 = [0,0], $serialNumber[], $serialAppearance[], $showValue[]
;Offsets
Global $bottomOffset = 10, $nextOffset = 15, $errorOffset = 40, $pickOffset = 5
;Checksums
Global $errorChecksum, $nextChecksum, $caseChecksum
;Rect values
Global $iX1, $iY1, $iX2, $iY2, $aPos, $sMsg, $sBMP_Path, $hRectangle_GUI
;Misc
Global $showVText = ""
;Debug
Global $debugTime = 200

HotKeySet("{HOME}", "TogglePause")
HotKeySet("{ESC}", "Terminate")

Initiate()

#cs 
	Initiate
		Introduces user to the program and prompts if user wants to load a preset.
		If user clicks yes then loads saved preset.
		Runs predrop to refresh page before product restock.
		If user clicks no then runs manual set up for all coordinates.
		Calibrates program values.
		Prompts user if they want to save current settings as a new preset.
		Loops until user hits HOME or begins program immediately if predrop was selected.
#ce
Func Initiate()
	Local $autoRefresh = 1
	Local $ready = MsgBox($MB_OKCANCEL, "Pop Bot", "This program will automatically click through pages of Pop Now figures and click the 'Pick One to Shake' button when one is available. There are some steps to ensure the program runs smoothly. Ready to set up the parameters for Pop Bot?")

	If ($ready = 1) Then
		Local $presetResponse = MsgBox($MB_YESNO, "Load Preset", "Do you want to load a preset?")
		If $presetResponse = 6 Then
			LoadPreset()
			$autoRefresh = PreDrop()
		ElseIf $presetResponse = 7 Then
			Setup()
		EndIf
	Else
		Terminate()
	EndIf
	If $autoRefresh = 1 Then ;If user said no during predop func, runs calibrate and asks if user wants to save preset
		Calibrate()
		SavePreset()

		MsgBox($MB_OK, "Pop Bot", "Set up complete! Press the 'Home' key to start and stop the program. Press the 'Escape' key to close or kill the program." & @CRLF & @CRLF & "Happy hunting!")

		While 1
			ToolTip("Pop Bot" & @CRLF & @CRLF & @CRLF & "Home to start" & @CRLF & "Escape to quit")
			Sleep(50)
		WEnd
	ElseIf $autoRefresh = 0 Then ;If user said yes, then  once predrop is complete, immediately run program
		TogglePause()
	EndIf
EndFunc ;== Initiate

#cs 
	Calibrate
		Moves the cursor to the PickButton location and waits so that accurate color capture is taken.
		Gets the correct color of the pick button since the pick button color changes when the cursor is on it.
#ce
Func Calibrate()
	Local $counter = 0
	Do
		MouseMove($pickButton[0], $pickButton[1])
		ToolTip("Calibrating " & $counter)
		;== Next button color
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
		$counter = $counter + 5
	Until $counter >= 100
	Tooltip("")
EndFunc ;== Calibrate

#cs 
	ShowParameters
		Moves the cursor between all defined coordinates for the different values
		Shows the colors that the program will look for when searching through boxes
#ce
Func ShowParameters()
	Local $mouseTemp = MouseGetPos()
	Local $nextParameter = 200
	Local $counterPreset = 0
	Local $num1 = 10
	Local $num2 = 20 ; error timer

	Do
		MouseMove($box1[0], $box1[1])
		ToolTip("Box position 1")
		$counterPreset = $counterPreset + $num1
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($box2[0], $box2[1])
		ToolTip("Box position 2")
		$counterPreset = $counterPreset + $num1
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($bottom[0], $bottom[1])
		ToolTip("Case position")
		$counterPreset = $counterPreset + $num1
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($error[0], $error[1])
		ToolTip("Error position")
		$counterPreset = $counterPreset + $num2
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($next[0], $next[1])
		ToolTip("Next button position")
		$counterPreset = $counterPreset + $num1
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($nextButton[0], $nextButton[1])
		ToolTip("Next click position")
		$counterPreset = $counterPreset + $num1
	Until $counterPreset >= $nextParameter
	ToolTip("")
	$counterPreset = 0

	Do
		MouseMove($pickButton[0], $pickButton[1])
		ToolTip("Pick button position")
		$counterPreset = $counterPreset + $num1
	Until $counterPreset >= $nextParameter

	ToolTip("")
	$counterPreset = 0

	ToolTip("")
	MouseMove($mouseTemp[0], $mouseTemp[1])

	GUICreate("", 200, 200)
    GUISetBkColor($iColorBox)
	$textRGB = _ColorGetRGB($iColorBox)

	$brightText = GUICtrlCreateLabel("Selected box color" & @CRLF & @CRLF & "Close this window to proceed", 10, 10)
	
	If $textRGB[0] < 128 or $textRGB[1] < 128 or $textRGB[2] < 128 Then
	GUICtrlSetColor($brightText, 0xFFFFFF)
	EndIf

    GUISetState(@SW_SHOW)

    While 1
            Switch GUIGetMsg()
                    Case $GUI_EVENT_CLOSE
                            ExitLoop
            EndSwitch
    WEnd

	GUIDelete()
EndFunc ;== ShowParameters

#cs 
	PreDrop
		Creates a GUI that prompts the user to enter a number for how long the program will wait until it automatically refreshes the page.
		The program waits the defined amount of time, makes sure the correct window is active, then sends F5 to refresh the page.
		Countdown restarts and continues until the Next Button is present.
		Once the next button is present the function exits and returns a value.
#ce
Func PreDrop()
	Local $refreshTimer = 1
	Local $nextPresent
	Local $confirm = MsgBox($MB_YESNO, "Pop Bot", "Do you want the bot to automatically refresh the page in anticipation for a Pop Now restock?")

	If $confirm = 6 Then
		Local $title = WinGetTitle("POP NOW", "")
		
		;refreshTimer creates GUI to get time amount.
		$refreshTimer = GetRefreshTimer()

		If $refreshTimer = "" or $refreshTimer = 0 Then
			$refreshTimer = 1
			MsgBox($MB_OK, "Pop Bot", "Default value of 1 minute has been set.")
		EndIf

		Local $refresh = $refreshTimer * 120000

		Do
			If WinExists($title) and $refresh < 1 Then
				Local $currentWin = WinGetTitle("[active]")
				Sleep(25)
				WinActivate($title)
				WinWaitActive($title)
				Send("{F5}")
				Sleep(25)
				WinActivate($currentWin)
				$refresh = $refreshTimer * 120000
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
		ToolTip("")
		Return 0
	ElseIf $confirm = 7 Then
		Return 1
	EndIf
EndFunc ;== PreDrop

#cs 
	GetRefreshTimer
		Creates GUI that prompts user to enter a number for how many minutes to wait until prorgam refreshes again.
#ce
Func GetRefreshTimer()
	$Form1 = GUICreate("Pop Bot", 301, 166, -1, -1, BitXOR($GUI_SS_DEFAULT_GUI, $WS_MINIMIZEBOX))
	GUICtrlCreateLabel("Ensure all windows are already in place if you are" & @CRLF & "planning to multitask while the bot refreshes the page." & @CRLF & @CRLF & "Enter number of minutes to wait in between each refresh:" & @CRLF & @CRLF & "Default value: 1", 10, 10)
	$Input1 = GUICtrlCreateInput("", 24, 100, 251, 21, $ES_NUMBER)
	GUICtrlSetLimit(-1, 3)
	$Button1 = GUICtrlCreateButton("OK", 44, 130, 75, 25)
	$Button2 = GUICtrlCreateButton("Cancel", 172, 130, 75, 25)
	GUISetState(@SW_SHOW)

	While 1
	    Switch GUIGetMsg()
	        Case $GUI_EVENT_CLOSE, $Button2
	            Exit
	        Case $Button1
	            $value = GUICtrlRead($Input1)
				GUIDelete($Form1)
				Return $value
	    EndSwitch
	WEnd
EndFunc

#cs 
	DupeCheck
		Takes a screenshot of area where the serial number is.
		Saves bitmap in an array.
		Takes screenshot of next serial number and compares to previous saved numbers.
		Records in a GUI the amount of duplicated cases there are.
#ce
Func DupeCheck()
	;Creates a screenshot path in temp folder
	Local $sImageFilePath = @TempDir & "\serialN.bmp"
	;captures screenshot of serial number area
	_ScreenCapture_Capture($sImageFilePath, $serial1[0], $serial1[1], $serial2[0], $serial2[1], False)
	;creates text file to store array data for serial number in temp folder
	Local $sFileName = @TempDir & "\serial.txt"
	Local $hFilehandle
	Local $newSN = True
	Local $tmpFile
	Local $tmpText
	;if its the first serial number it clears the temp files
	;else it adds to the temp file
	If $freshStart = true Then 
		$hFilehandle = FileOpen($sFileName, $FO_OVERWRITE)
	Else
		$hFilehandle = FileOpen($sFileName, $FO_APPEND)
	EndIf
	;detects text, specifically numbers from screenshot and only saves the last 10 numbers
	;sometimes the multiple 0's turns into a space and i dont know why, but it breaks everything so instead i save the 10 numbers at the end
	Local $sOCRTextResult = _UWPOCR_GetText($sImageFilePath)
	Local $str = StringRegExpReplace($sOCRTextResult, "\D", "")
	Local $num = StringRight($str, 10)
	Local $nbr = Number($num)
	Local $counter = $serialCounter - 1
	$serialNumber[$serialCounter] = $nbr
	$serialAppearance[$serialCounter] = 1
	;after the first serial number, it compares the current number to the previous ones for repeats.
	If $freshStart = false Then
		Do
			If $serialNumber[$serialCounter] = $serialNumber[$counter] Then
				$serialAppearance[$counter] = $serialAppearance[$counter] + 1 ;adds 1 to the number of appearances a serial number has shown
				$newSN = False
				ExitLoop
			Else
				$counter = $counter - 1 ;goes to the next number in the list
			EndIf
		Until $counter < 0
	EndIf

	If $newSN = True Then ;new serial number gets added to the array
		FileWriteLine($hFilehandle, $serialNumber[$serialCounter])
		FileWriteLine($hFilehandle, $serialAppearance[$serialCounter])
		$uniquePassed = $uniquePassed + 1
	ElseIf $newSN = False Then ;repeated serial number does not get added and the number of appearances gets updated in txt file
		$tmpFile = @TempDir & "\serial.txt"
		$tmpText = FileRead(FileGetLongName($tmpFile),FileGetSize(FileGetLongName($tmpFile)))
		$tmpText = StringReplace($tmpText, $serialNumber[$counter] & @CRLF & $serialAppearance[$counter] - 1, $serialNumber[$counter] & @CRLF & $serialAppearance[$counter])
		FileDelete($tmpFile)
		FileWrite($tmpFile, $tmpText)
		$dupePassed = $dupePassed + 1
	EndIf
	FileClose($sFileName)
	$serialCounter = $serialCounter + 1
	$freshStart = false
EndFunc

#cs 
	TogglePause
		Uses values from setup or loadpreset to check if the case of boxes is in the correct position.
		If true, then checks if the next button is not currently loading.
		If true, then program searches the area of the case for the color the user defined as an "available box".
		If true then clicks Pick One button to attempt to buy a box.
		If clicked then program checks if an error message appears at the top of the screen (meaning the box is no longer available).
		If error is there then program waits a little bit and then resumes the search.
		If error is not there then program plays a success sound which can be replaced manually.
		(Success sound can be replaced manually.)
#ce
Func TogglePause()
	$g_bPaused = Not $g_bPaused

	Local $title
	Local $successVolume = 10

	;If box was already found, but user wants to go back and continue, this will reset value counters and stop the success sound
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
		ToolTip("Pop Bot - PAUSED" & @CRLF & @CRLF & $casesPassed & " total cases passed" & @CRLF & $uniquePassed & " unique cases" & @CRLF & $dupePassed & " duplicate cases" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
		Sleep(50)
	WEnd ;== Program paused

	;== Program unpaused
	While $g_bPaused = 1
		ToolTip("Pop Bot" & @CRLF & @CRLF & $casesPassed & " total cases passed" & @CRLF & $uniquePassed & " unique cases" & @CRLF & $dupePassed & " duplicate cases" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
		Sleep(50)

		;Checks if the Pick button is white which most likely means the page timed out
		While PixelGetColor($pickButton[0], $pickButton[1]) = $iColorWhite and $timeoutTimer > 0
			ToolTip("Pop Bot" & @CRLF & Floor($timeoutTimer / 20) + 1 & " until timeout refresh")
			$timeoutTimer = $timeoutTimer - 1
			Sleep(50)
		WEnd

		If $timeoutTimer < 1 and WinGetTitle("[ACTIVE]") = $title Then
			Send("{F5}")
			$timeoutTimer = $timeoutReset + 200
			Sleep(50)
		ElseIf $timeoutTimer < 1 and WinGetTitle("[ACTIVE]") = not $title Then
			Send("!{LEFT}")
			$timeoutTimer = $timeoutReset
			Sleep(50)
		Else
			$timeoutTimer = $timeoutReset
			Sleep(50)
		EndIf

		;== Next check
		If ($nextChecksum = PixelChecksum($next[0] - 15, $next[1] - 15, $next[0] + 15, $next[1] + 15)) Then
			;== Case check
			If ($caseChecksum = PixelChecksum($bottom[0] - $bottomOffset, $bottom[1] - $bottomOffset, $bottom[0] + $bottomOffset, $bottom[1] + $bottomOffset)) Then
				Sleep(50)
				;Checks if case is a duplicate
				DupeCheck()

				$boxes = PixelSearch($box1[0], $box1[1], $box2[0], $box2[1], $iColorBox, 0)
				;== Box check
				If IsArray($boxes) Then
					;== Mouse clicks pick button
					MouseMove($pickButton[0], $pickButton[1], 0)
					MouseClick($MOUSE_CLICK_LEFT)

					ToolTip("Pop Bot" & @CRLF & @CRLF & $casesPassed & " total cases passed" & @CRLF & $uniquePassed & " unique cases" & @CRLF & $dupePassed & " duplicate cases" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
					Sleep(50)

					Local $errorPresent = 0
					Local $errorTimer = 0
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

						ToolTip("Pop Bot" & @CRLF & @CRLF & $casesPassed & " total cases passed" & @CRLF & $uniquePassed & " unique cases" & @CRLF & $dupePassed & " duplicate cases" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)

						Sleep(1000)
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
								Local $tSearch = FileFindFirstFile("Success Sound\*.wav")
								Local $sFileName = FileFindNextFile($tSearch)
                                SoundSetWaveVolume($successVolume)
							    SoundPlay("Success Sound\" & $sFileName)
							EndIf
                            Sleep(50)
						WEnd
					EndIf
				;== Box unavailable
				Else
					MouseMove($nextButton[0], $nextButton[1], 0)

					ToolTip("Pop Bot" & @CRLF & @CRLF & $casesPassed & " total cases passed" & @CRLF & $uniquePassed & " unique cases" & @CRLF & $dupePassed & " duplicate cases" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
					Sleep(50)

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
			EndIf ;== Case check
			$boxes = 0
		EndIf ;== Next check
		$boxes = 0
	WEnd ;== Program unpaused
	ToolTip("")
EndFunc ;== TogglePause

#cs 
	ShowValue
		Func shows a red box outline of where user selected for set up value
#ce
Func ShowValue()
	;Shows a redbox outline, top line, left line, bottom line, right line
	$popupT = GUICreate("", $showValue[0], $showValue[1], $showValue[2], $showValue[3], $WS_POPUP, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
	GUISetBkColor($iColorRed, $popupT)
	$popupL = GUICreate("", $showValue[4], $showValue[5], $showValue[6], $showValue[7], $WS_POPUP, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
	GUISetBkColor($iColorRed, $popupL)
	$popupB = GUICreate("", $showValue[8], $showValue[9], $showValue[10], $showValue[11], $WS_POPUP, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
	GUISetBkColor($iColorRed, $popupB)
	$popupR = GUICreate("", $showValue[12], $showValue[13], $showValue[14], $showValue[15], $WS_POPUP, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
	GUISetBkColor($iColorRed, $popupR)
	$prompt = GUICreate("Setup", 245, 60, $showValue[6], $showValue[11] + 5)
	GUICtrlCreateLabel($showVText, 10, 10)
	Local $cancelButton	= GUICtrlCreateButton("Cancel", 7, 30, 70, 25)
	Local $againButton	= GUICtrlCreateButton("Try Again", 87, 30, 70, 25)
	Local $contButton	= GUICtrlCreateButton("Continue", 167, 30, 70, 25)
	GUISetState(@SW_SHOW, $popupT)
	GUISetState(@SW_SHOW, $popupL)
	GUISetState(@SW_SHOW, $popupB)
	GUISetState(@SW_SHOW, $popupR)
	;show continue, try again, cancel pop up to ask if red box is correct
	GUISetState(@SW_SHOW, $prompt)
	WinActivate("Setup")
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE, $cancelButton
				Terminate()
			Case $contButton
				GUIDelete($popupT)
				GUIDelete($popupL)
				GUIDelete($popupB)
				GUIDelete($popupR)
				GUIDelete($prompt)
				Return 11
			Case $againButton
				GUIDelete($popupT)
				GUIDelete($popupL)
				GUIDelete($popupB)
				GUIDelete($popupR)
				GUIDelete($prompt)
				Return 10
		EndSwitch
	WEnd
EndFunc ;== ShowValue

#cs 
	GetBoxPos
		Asks user draw a rectangle around area of boxes on the screen
		calls mark_rect
		calls show value to show user
#ce
Func GetBoxPos()
	Local $response
	Do
		Local $hMain_GUI_B 		= GUICreate("Setup", 300, 100)
		Local $hRect_Button 	= GUICtrlCreateButton("Mark Area",  30, 50, 80, 30)
		Local $hCancel_Button 	= GUICtrlCreateButton("Cancel", 180, 50, 80, 30)
		GUICtrlCreateLabel("Draw a box around the location of the figure boxes.", 25, 20)
		GUISetState()
		While 1
		    Switch GUIGetMsg()
		        Case $GUI_EVENT_CLOSE, $hCancel_Button
		            FileDelete(@TempDir & "\Rect.bmp")
					Terminate()
		        Case $hRect_Button
		            GUIDelete($hMain_GUI_B)
		            Mark_Rect()
					$box1[0] = $iX1
					$box1[1] = $iY1
					$box2[0] = $iX2
					$box2[1] = $iY2
					;Top line
					$showValue[0]	= $iX2 - $iX1	;width
					$showValue[1]	= 1				;height
					$showValue[2]	= $iX1			;left
					$showValue[3]	= $iY1			;top
					;Left line
					$showValue[4]	= 1				;width
					$showValue[5]	= $iY2 - $iY1	;height
					$showValue[6]	= $iX1			;left, prompt left
					$showValue[7]	= $iY1			;top
					;Bottom line
					$showValue[8]	= $iX2 - $iX1	;width
					$showValue[9]	= 1				;height
					$showValue[10]	= $iX1			;left
					$showValue[11]	= $iY2			;top, prompt top
					;Right line
					$showValue[12]	= 1				;width
					$showValue[13]	= $iY2 - $iY1	;height
					$showValue[14]	= $iX2			;left
					$showValue[15]	= $iY1			;top
					;Prompt text
					$showVText		= "Selected area to search for boxes."
					;Confirm selection with user
					$response = ShowValue()
					ExitLoop
		    EndSwitch
		WEnd
		$caseLeft[0] = $box1[0]
		$caseLeft[1] = $box1[1]
	Until $response = 11
EndFunc ;== GetBoxPos

#cs 
	GetCasePos
		asks user to press ctrl key on the bottom of the case of boxes
		calls showvalue to confirm location
#ce
Func GetCasePos()
	Local $response
	Local $UserDLL = DllOpen("user32.dll")
	Do
		While 1
			ToolTip("Press 'CTRL' on the bottom of the case")
			If _IsPressed("11", $UserDLL) Then 
				$bottom = MouseGetPos()
				ToolTip("")
				ExitLoop
			EndIf
			Sleep(50)
		Wend
		;Top line
		$showValue[0]	= $bottomOffset						;width
		$showValue[1]	= 1									;height
		$showValue[2]	= $bottom[0] - ($bottomOffset / 2)	;left
		$showValue[3]	= $bottom[1] - ($bottomOffset / 2)	;top
		;Left line
		$showValue[4]	= 1									;width
		$showValue[5]	= $bottomOffset						;height
		$showValue[6]	= $bottom[0] - ($bottomOffset / 2)	;left, prompt left
		$showValue[7]	= $bottom[1] - ($bottomOffset / 2)	;top
		;Bottom line
		$showValue[8]	= $bottomOffset + 1					;width
		$showValue[9]	= 1									;height
		$showValue[10]	= $bottom[0] - ($bottomOffset / 2)	;left
		$showValue[11]	= $bottom[1] + ($bottomOffset / 2)	;top, prompt top
		;Right line
		$showValue[12]	= 1									;width
		$showValue[13]	= $bottomOffset						;height
		$showValue[14]	= $bottom[0] + ($bottomOffset / 2)	;left
		$showValue[15]	= $bottom[1] - ($bottomOffset / 2)	;top
		;Prompt text
		$showVText		= "Bottom of case position."
		;Confirm selection with user
		$response = ShowValue()
	Until $response = 11
	DllClose($UserDLL)
EndFunc ;== GetCasePos

#cs 
	GetSerialPos
		asks user to press ctrl key on the serial number
		from selected point, func finds where black text is located
		checks left and right for more black text, once where is enough non-black space, marks left and right bounds
		checks upper and lower to find top and bottom bounds
		calls showvalue to confirm location
#ce
Func GetSerialPos()
	Local $response
	Local $tempPos
	Local $counter = 1
	Local $i
	Local $UserDLL = DllOpen("user32.dll")
	Do
		While 1
			ToolTip("Press 'CTRL' on the case serial number")
			If _IsPressed("11", $UserDLL) Then 
				$tempPos = MouseGetPos()
				ToolTip("")
				ExitLoop
			EndIf
			Sleep(50)
		WEnd
		While 1
			$i = PixelSearch($tempPos[0], $tempPos[1], $tempPos[0], $tempPos[1], $iColorBlack, 200)
			If @error Then
				$xL = PixelSearch($tempPos[0] - $counter, $tempPos[1], $tempPos[0] - $counter, $tempPos[1], $iColorBlack, 200)
				If @error Then
					$xTL = PixelSearch($tempPos[0] - $counter, $tempPos[1] - $counter, $tempPos[0] - $counter, $tempPos[1] - $counter, $iColorBlack, 200)
					If @error Then
						$xT = PixelSearch($tempPos[0], $tempPos[1] - $counter, $tempPos[0], $tempPos[1] - $counter, $iColorBlack, 200)
						If @error Then
							$xTR = PixelSearch($tempPos[0] + $counter, $tempPos[1] - $counter, $tempPos[0] + $counter, $tempPos[1] - $counter, $iColorBlack, 200)
							If @error Then
								$xR = PixelSearch($tempPos[0] + $counter, $tempPos[1], $tempPos[0] + $counter, $tempPos[1], $iColorBlack, 200)
								If @error Then
									$xBR = PixelSearch($tempPos[0] + $counter, $tempPos[1] + $counter, $tempPos[0] + $counter, $tempPos[1] + $counter, $iColorBlack, 200)
									If @error Then
										$xB = PixelSearch($tempPos[0], $tempPos[1] + $counter, $tempPos[0], $tempPos[1] + $counter, $iColorBlack, 200)
										If @error Then
											$xBL = PixelSearch($tempPos[0] - $counter, $tempPos[1] + $counter, $tempPos[0] - $counter, $tempPos[1] + $counter, $iColorBlack, 200)
											If @error Then
												$counter = $counter + 1
												If $counter > 50 Then MsgBox(0, "", "Serial Number not found")
											Else
												$tempPos = $xBL
												$tempPos[1] = $tempPos[1] + 2
												ExitLoop
											EndIf
										Else
											$tempPos = $xB
											$tempPos[1] = $tempPos[1] + 2
											ExitLoop
										EndIf
									Else
										$tempPos = $xBR
										$tempPos[1] = $tempPos[1] + 2
										ExitLoop
									EndIf
								Else
									$tempPos = $xR
									ExitLoop
								EndIf
							Else
								$tempPos = $xTR
								$tempPos[1] = $tempPos[1] - 2
								ExitLoop
							EndIf
						Else
							$tempPos = $xT
							$tempPos[1] = $tempPos[1] - 2
							ExitLoop
						EndIf
					Else
						$tempPos = $xTL
						$tempPos[1] = $tempPos[1] - 2
						ExitLoop
					EndIf
				Else
					$tempPos = $xL
					ExitLoop
				EndIf
			Else
				$tempPos = $i
				ExitLoop
			EndIf
		WEnd
		Local $boundL = 0
		Local $boundR = 0
		Local $counterL = 0
		Local $counterR = 0
		While 1
			If $boundL < 30 Then
				Local $iSN = PixelSearch($tempPos[0] - $counterL, $tempPos[1], $tempPos[0] - $counterL, $tempPos[1], $iColorBlack, 200)
				If @error Then
					$counterL = $counterL + 1
					$boundL = $boundL + 1
				Else
					$counterL = $counterL + 1
					$boundL = 0
				EndIf
			EndIf
			If $boundR < 30 Then
				Local $jSN = PixelSearch($tempPos[0] + $counterR, $tempPos[1], $tempPos[0] + $counterR, $tempPos[1], $iColorBlack, 200)
				If @error Then
					$counterR = $counterR + 1
					$boundR = $boundR + 1
				Else
					$counterR = $counterR + 1
					$boundR = 0
				EndIf
			EndIf
			If $boundL > 29 and $boundR > 29 Then ExitLoop
		WEnd

		$serial1[0] = $tempPos[0] - $counterL
		$serial2[0] = $tempPos[0] + $counterR

		Local $boundT = 0
		Local $boundB = 0
		Local $counterT = 0
		Local $counterB = 0
		$counterL = 0
		$counterR = 0

		While 1
			If $boundT < 160 Then
				Local $kSN = PixelSearch($tempPos[0] - $counterL, $tempPos[1] - $counterT, $tempPos[0] - $counterL, $tempPos[1] - $counterT, $iColorBlack, 200)
				If @error and $counterL < 8 Then
					$counterL = $counterL + 1
					$boundT = $boundT + 1
				ElseIf $counterL > 7 Then
					$counterL = 0
					$counterT = $counterT + 1
				Else
					$counterL = 0
					$counterT = $counterT + 1
					$boundT = 0
				EndIf
			EndIf
			If $boundB < 160 Then
				Local $lSN = PixelSearch($tempPos[0] + $counterR, $tempPos[1] + $counterB, $tempPos[0] + $counterR, $tempPos[1] + $counterB, $iColorBlack, 200)
				If @error and $counterR < 8 Then
					$counterR = $counterR + 1
					$boundB = $boundB + 1
				ElseIf $counterR > 7 Then
					$counterR = 0
					$counterB = $counterB + 1
				Else
					$counterR = 0
					$counterB = $counterB + 1
					$boundB = 0
				EndIf
			EndIf
			If $boundT > 159 and $boundB > 159 Then ExitLoop
		WEnd

		$serial1[1] = $tempPos[1] - $counterT
		$serial2[1] = $tempPos[1] + $counterB

		;Top line
		$showValue[0]	= $serial2[0] - $serial1[0]	;width
		$showValue[1]	= 1							;height
		$showValue[2]	= $serial1[0]				;left
		$showValue[3]	= $serial1[1]				;top
		;Left line
		$showValue[4]	= 1							;width
		$showValue[5]	= $serial2[1] - $serial1[1]	;height
		$showValue[6]	= $serial1[0]				;left, prompt left
		$showValue[7]	= $serial1[1]				;top
		;bottom line
		$showValue[8]	= $serial2[0] - $serial1[0]	;width
		$showValue[9]	= 1							;height
		$showValue[10]	= $serial1[0]				;left
		$showValue[11]	= $serial2[1]				;top, prompt top
		;Right line
		$showValue[12]	= 1							;width
		$showValue[13]	= $serial2[1] - $serial1[1]	;height
		$showValue[14]	= $serial2[0]				;left
		$showValue[15]	= $serial1[1]				;top
		;Prompt text
		$showVText		= "Case serial number position."
		;Confirm selection with user
		$response = ShowValue()

		If $response = 11 Then
			$sBMP_Path_SN = @TempDir & "\serialN.bmp"
			_ScreenCapture_Capture($sBMP_Path_SN, $serial1[0], $serial1[1], $serial2[0], $serial2[1], False)
			
			Local $serialTest = _UWPOCR_GetText($sBMP_Path_SN)
			Local $nbr = StringRight($serialTest, 10)
			;MsgBox(0, "", $serialTest)
			;MsgBox(0, "", $nbr)
			If $serialTest = "" or StringLen($nbr) <> 10 Then
				If MsgBox($MB_OKCANCEL, "Error", "Serial number illegible or incorrect number of digits. Try again.") = 2 Then Terminate()
				$response = 10
			EndIf
		EndIf
	Until $response = 11
	DllClose($UserDLL)
EndFunc ;== GetSerialPos

Func GetErrorPos()
	Local $response
	Local $UserDLL = DllOpen("user32.dll")
	Do
		While 1
			ToolTip("Press 'CTRL' in the center of the banner" & @CRLF & "above the 'Shake for Hints' text")
			If _IsPressed("11", $UserDLL) Then
				$error = MouseGetPos()
				ToolTip("")
				ExitLoop
			EndIf
			Sleep(50)
		WEnd
		;Top line
		$showValue[0]	= $errorOffset						;width
		$showValue[1]	= 1									;height
		$showValue[2]	= $error[0] - ($errorOffset / 2)	;left
		$showValue[3]	= $error[1] - ($errorOffset / 2)	;top
		;Left line
		$showValue[4]	= 1									;width
		$showValue[5]	= $errorOffset						;height
		$showValue[6]	= $error[0] - ($errorOffset / 2)	;left, prompt left
		$showValue[7]	= $error[1] - ($errorOffset / 2)	;top
		;bottom line
		$showValue[8]	= $errorOffset + 1					;width
		$showValue[9]	= 1									;height
		$showValue[10]	= $error[0] - ($errorOffset / 2)	;left
		$showValue[11]	= $error[1] + ($errorOffset / 2)	;top, prompt top
		;Right line
		$showValue[12]	= 1									;width
		$showValue[13]	= $errorOffset						;height
		$showValue[14]	= $error[0] + ($errorOffset / 2)	;left
		$showValue[15]	= $error[1] - ($errorOffset / 2)	;top
		;Prompt text
		$showVText		= "Estimated error message position."
		;Confirm selection with user
		$response = ShowValue()
	Until $response = 11
	DllClose($UserDLL)
EndFunc ;== GetErrorPos

Func GetNextPos()
	Local $response
	Local $UserDLL = DllOpen("user32.dll")
	Do
		While 1
			ToolTip("Press 'CTRL' in the center of next case button")
			If _IsPressed("11", $UserDLL) Then
				$next = MouseGetPos()
				$nextButton[0] = $next[0]
				$nextButton[1] = $next[1] + $nextOffset + 2
				ToolTip("")
				ExitLoop
			EndIf
			Sleep(50)
		WEnd
		;Top line
		$showValue[0]	= $nextOffset						;width
		$showValue[1]	= 1									;height
		$showValue[2]	= $next[0] - ($nextOffset / 2)	;left
		$showValue[3]	= $next[1] - ($nextOffset / 2)	;top
		;Left line
		$showValue[4]	= 1									;width
		$showValue[5]	= $nextOffset						;height
		$showValue[6]	= $next[0] - ($nextOffset / 2)	;left, prompt left
		$showValue[7]	= $next[1] - ($nextOffset / 2)	;top
		;bottom line
		$showValue[8]	= $nextOffset + 1					;width
		$showValue[9]	= 1									;height
		$showValue[10]	= $next[0] - ($nextOffset / 2)	;left
		$showValue[11]	= $next[1] + ($nextOffset / 2)	;top, prompt top
		;Right line
		$showValue[12]	= 1									;width
		$showValue[13]	= $nextOffset						;height
		$showValue[14]	= $next[0] + ($nextOffset / 2)	;left
		$showValue[15]	= $next[1] - ($nextOffset / 2)	;top
		;Prompt text
		$showVText		= "Next button position."
		;Confirm selection with user
		$response = ShowValue()
	Until $response = 11
	DllClose($UserDLL)
EndFunc ;== GetNextPos

Func GetPickPos()
	Local $response
	Local $UserDLL = DllOpen("user32.dll")
	Do
		While 1
			ToolTip("Press 'CTRL' on the 'Pick One to Shake' button")
			If _IsPressed("11", $UserDLL) Then
				$pickButton = MouseGetPos()
				ToolTip("")
				ExitLoop
			EndIf
			Sleep(50)
		WEnd
		;Top line
		$showValue[0]	= $pickOffset							;width
		$showValue[1]	= 1										;height
		$showValue[2]	= $pickButton[0] - ($pickOffset / 2)	;left
		$showValue[3]	= $pickButton[1] - ($pickOffset / 2)	;top
		;Left line
		$showValue[4]	= 1										;width
		$showValue[5]	= $pickOffset							;height
		$showValue[6]	= $pickButton[0] - ($pickOffset / 2)	;left, prompt left
		$showValue[7]	= $pickButton[1] - ($pickOffset / 2)	;top
		;bottom line
		$showValue[8]	= $pickOffset + 1						;width
		$showValue[9]	= 1										;height
		$showValue[10]	= $pickButton[0] - ($pickOffset / 2)	;left
		$showValue[11]	= $pickButton[1] + ($pickOffset / 2)	;top, prompt top
		;Right line
		$showValue[12]	= 1										;width
		$showValue[13]	= $pickOffset							;height
		$showValue[14]	= $pickButton[0] + ($pickOffset / 2)	;left
		$showValue[15]	= $pickButton[1] - ($pickOffset / 2)	;top
		;Prompt text
		$showVText		= "Pick button position."
		;Confirm selection with user
		$response = ShowValue()
	Until $response = 11
	DllClose($UserDLL)
EndFunc ;== GetPickPos

Func GetBoxColor()
	Local $response
	Local $UserDLL = DllOpen("user32.dll")
	Do
		While 1
			ToolTip("Press 'CTRL' on a unique color of a figure box")
			If _IsPressed("11", $UserDLL) Then
				$boxPos = MouseGetPos()
				ToolTip("")
				ExitLoop
			EndIf
			Sleep(50)
		WEnd
	
		GUICreate("Setup", 200, 200) ; will create a dialog box that when displayed is centered
		$colorCase = "0x" & Hex(PixelGetColor($boxPos[0], $boxPos[1]), 6)
    	GUISetBkColor($colorCase)
		$textRGB = _ColorGetRGB($colorCase)
		$brightText = GUICtrlCreateLabel("Selected box color", 10, 10)
		If $textRGB[0] < 128 or $textRGB[1] < 128 or $textRGB[2] < 128 Then GUICtrlSetColor($brightText, 0xFFFFFF)
    	GUISetState(@SW_SHOW)

    	; Loop until the user exits.
    	Local $cancelButtonC 	= GUICtrlCreateButton("Cancel", 10, 160, 60, 25)
		Local $againButtonC 	= GUICtrlCreateButton("Try Again", 70, 160, 60, 25)
		Local $contButtonC 		= GUICtrlCreateButton("Continue", 130, 160, 60, 25)
		GUISetState()
		WinActivate("Setup")

		While 1
			Switch GUIGetMsg()
				Case $GUI_EVENT_CLOSE, $cancelButtonC
					Terminate()
				Case $contButtonC
					$response = 11
					GUIDelete()
					ExitLoop
				Case $againButtonC
					$response = 10
					GUIDelete()
					ExitLoop
			EndSwitch
		WEnd

		GUIDelete()
	
		$iColorBox = PixelGetColor($boxPos[0], $boxPos[1])

	Until $response = 11
	DllClose($UserDLL)
EndFunc ;== GetBoxColor

#cs 
	Setup
		Calls each function to get values on corresponding places for the program to click.
#ce
Func Setup()
	GetBoxPos() 
	GetCasePos()
	GetSerialPos()
	GetErrorPos()
	GetNextPos()
	GetPickPos()
	GetBoxColor()
EndFunc ;== Setup

#cs 
	LoadPreset
		Creates a GUI and reads all titles of txt files in Presets folder.
		User is able to click on a preset and either load it or delete it.
		If user clicks load, then values in txt file is saved as values in program and showParameters is called.
		If user clicks delete, prompts to make sure user wants to delete, then deletes txt file and refreshes GUI.
#ce
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
			Terminate()
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
	Local $lineCounter = 0
	Do
		$lineCounter = $lineCounter + 1
		FileReadLine($rFileHandle, $lineCounter)
	Until FileReadLine($rFileHandle, $lineCounter) = ""
	
	If $lineCounter = 16 Then
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
	ElseIf $lineCounter = 26 Then
		$box1[0]		= FileReadLine($rFileHandle, 1)
		$box1[1]		= FileReadLine($rFileHandle, 2)
		$box2[0]		= FileReadLine($rFileHandle, 3)
		$box2[1]		= FileReadLine($rFileHandle, 4)
		$bottom[0]		= FileReadLine($rFileHandle, 5)
		$bottom[1]		= FileReadLine($rFileHandle, 6)
		$caseChecksum	= FileReadLine($rFileHandle, 7)
		$iColorCase		= FileReadLine($rFileHandle, 8)
		$iColorCaseLeft	= FileReadLine($rFileHandle, 9)
		$error[0]		= FileReadLine($rFileHandle, 10)
		$error[1]		= FileReadLine($rFileHandle, 11)
		$errorChecksum	= FileReadLine($rFileHandle, 12)
		$next[0]		= FileReadLine($rFileHandle, 13)
		$next[1]		= FileReadLine($rFileHandle, 14)
		$nextChecksum	= FileReadLine($rFileHandle, 15)
		$nextButton[0]	= FileReadLine($rFileHandle, 16)
		$nextButton[1]	= FileReadLine($rFileHandle, 17)
		$pickButton[0]	= FileReadLine($rFileHandle, 18)
		$pickButton[1]	= FileReadLine($rFileHandle, 19)
		$iColorPick		= FileReadLine($rFileHandle, 20)
		$iColorBox		= FileReadLine($rFileHandle, 21)
		$serial1[0]		= FileReadLine($rFileHandle, 22)
		$serial1[1]		= FileReadLine($rFileHandle, 23)
		$serial2[0]		= FileReadLine($rFileHandle, 24)
		$serial2[1]		= FileReadLine($rFileHandle, 25)
	Else
		FileClose($rFileHandle)
		MsgBox($MB_SYSTEMMODAL, "Load Preset", "Error loading preset." & @CRLF & @CRLF & "Application closing")
		Terminate()
	EndIf

	FileClose($rFileHandle)

	;ShowParameters()
EndFunc ;== LoadPreset

#cs 
	SavePreset
		After LoadPreset or Setup, prompts user to save preset.
		If user saves then create input box to name the preset.
		Preset is saved as a txt file in the preset folder.
		If file is unable to be created then prorgam closes.
#ce
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
	FileWriteLine($hFileHandle, $box1[0])
	FileWriteLine($hFileHandle, $box1[1])
	FileWriteLine($hFileHandle, $box2[0])
	FileWriteLine($hFileHandle, $box2[1])
	FileWriteLine($hFileHandle, $bottom[0])
	FileWriteLine($hFileHandle, $bottom[1])
	FileWriteLine($hFileHandle, $caseChecksum)
	FileWriteLine($hFileHandle, $iColorCase)
	FileWriteLine($hFileHandle, $iColorCaseLeft)
	FileWriteLine($hFileHandle, $error[0])
	FileWriteLine($hFileHandle, $error[1])
	FileWriteLine($hFileHandle, $errorChecksum)
	FileWriteLine($hFileHandle, $next[0])
	FileWriteLine($hFileHandle, $next[1])
	FileWriteLine($hFileHandle, $nextChecksum)
	FileWriteLine($hFileHandle, $nextButton[0])
	FileWriteLine($hFileHandle, $nextButton[1])
	FileWriteLine($hFileHandle, $pickButton[0])
	FileWriteLine($hFileHandle, $pickButton[1])
	FileWriteLine($hFileHandle, $iColorPick)
	FileWriteLine($hFileHandle, $iColorBox)
	FileWriteLine($hFileHandle, $serial1[0])
	FileWriteLine($hFileHandle, $serial1[1])
	FileWriteLine($hFileHandle, $serial2[0])
	FileWriteLine($hFileHandle, $serial2[1])

    ; Close the handle returned by FileOpen.
    FileClose($sFileName)

	MsgBox($MB_OK, "Save Preset", "Preset saved." & @CRLF & "File location: " & FileGetLongName($sFileName))
EndFunc ;== SavePreset

#cs 
	Func Mark_Rect
		code was taken from Melba23 on autoitscript forums.
		https://www.autoitscript.com/forum/topic/95920-solved-selecting-a-rectangle-to-screencapture/
		added tooltip messages to assist in guiding user.
#ce
Func Mark_Rect()
    Local $aMouse_Pos, $aMask, $aM_Mask, $iTemp
    Local $UserDLL = DllOpen("user32.dll")

	; Wait until mouse button pressed
    While Not _IsPressed("01", $UserDLL)
		ToolTip("Pop Bot" & @CRLF & @CRLF & "Click and Drag to select area")
        Sleep(50)
    WEnd

	Local $aPos, $aData = _WinAPI_EnumDisplayMonitors()

	If IsArray($aData) Then
	        ReDim $aData[$aData[0][0] + 1][5]
	        For $i = 1 To $aData[0][0]
	                $aPos = _WinAPI_GetPosFromRect($aData[$i][1])
	                For $j = 0 To 3
	                        $aData[$i][$j + 1] = $aPos[$j]
	                Next
	        Next
	EndIf

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

#cs 
	Func Terminate
		calls Exit to close program.
#ce
Func Terminate()
	Exit
EndFunc ;== Terminate