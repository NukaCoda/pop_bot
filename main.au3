;version 19
#include <WinAPIGdi.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>
#include <ScreenCapture.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Color.au3>
#Include <Misc.au3>

Global $g_bPaused = False, $iColorBox = 0xFF00FF, $iColorBottom = 0xFF00FF, $iColorNext = 0x000000, $mousePos = [0,0]
Global $box1 = [0,0], $box2 = [0,0], $next = [0,0], $bottom = [0,0], $pickButton = [0,0], $nextButton = [0,0], $error = [0,0], $caseLeft = [0, 0]
Global $iTimeout = 10
Global $scale = _WinAPI_EnumDisplaySettings('', $ENUM_CURRENT_SETTINGS)[0] / @DesktopWidth
Global $casesPassed = 0, $resumeAfterFind = 0
Global $errorChecksum, $nextChecksum, $caseChecksum, $caseColor
Global $debugTime = 200
Global $iX1, $iY1, $iX2, $iY2, $aPos, $sMsg, $sBMP_Path

HotKeySet("{HOME}", "TogglePause")
HotKeySet("{ESC}", "Terminate")

$ready = MsgBox($MB_OKCANCEL, "Pop Now Bot", "This program will automatically click through pages of Pop Now figures and click the 'Pick One to Shake' button when one is available. There are some steps to ensure the program runs smoothly. Ready to set up the parameters for Pop Now Bot?")

If ($ready = 1) Then
	$presetResponse = MsgBox($MB_YESNO, "Load Preset", "Do you want to load a preset?")
	If $presetResponse = 6 Then
		LoadPreset()
	ElseIf $presetResponse = 7 Then
		Setup()
	EndIf
Else
	Terminate()
EndIf

Sleep(100)

$mouseTemp = MouseGetPos()
MouseMove(0, 0, 0)

Sleep(10)

;== Error message location
$errorChecksum = PixelChecksum($error[0] - 10, $error[1] - 10, $error[0] + 10, $error[1] + 10)
;== Next loaded location
$nextChecksum = PixelChecksum($next[0] - 15, $next[1] - 15, $next[0] + 15, $next[1] + 15)
;== Case in position location
$caseChecksum = PixelChecksum($bottom[0] - 1, $bottom[1] - 1, $bottom[0] + 1, $bottom[1] + 1)
$caseColor = PixelGetColor($bottom[0], $bottom[1])
;== Left of case color
$caseLeftColor = PixelGetColor($caseLeft[0], $caseLeft[1])

Sleep(10)

MouseMove($mouseTemp[0], $mouseTemp[1], 0)

;== Save Preset
If (MsgBox($MB_YESNO, "Save Preset", "Would you like to save these parameters as a preset?") = 6) Then
	SavePreset()
EndIf

MsgBox($MB_OK, "Pop Now Bot", "Set up complete! Press the 'Home' key to start and stop the program. Press the 'Escape' key to close or kill the program." & @CRLF & @CRLF & "Happy hunting!")

While 1
	ToolTip("Pop Now Bot" & @CRLF & @CRLF & @CRLF & "Home to start" & @CRLF & "Escape to quit")
			Sleep(50)
WEnd
	
Func TogglePause()
	$g_bPaused = Not $g_bPaused
	If $resumeAfterFind = 1 Then
		$casesPassed = 0
		$resumeAfterFind = 0
	EndIf
	;== Program paused
	While $g_bPaused = 0
		If $casesPassed = 1 Then
			ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " case passed - PAUSED" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
			Sleep(50)
		Else
			ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " cases passed - PAUSED" & @CRLF & @CRLF & "Home to resume" & @CRLF & "Escape to quit")
			Sleep(50)
		EndIf
	WEnd ;== Program paused

	;== Program unpaused
	While $g_bPaused = 1
		If $casesPassed = 1 Then
			ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " case passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
			Sleep(50)
		Else
			ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " cases passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
			Sleep(50)
		EndIf

		;== Next check
		;== Next button loaded
		;ToolTip("Next check")
		;Sleep(100)
		If ($nextChecksum = PixelChecksum($next[0] - 15, $next[1] - 15, $next[0] + 15, $next[1] + 15)) Then
			;== Case check
			;== Case in position
			;ToolTip("Case check")
			;Sleep(100)
			If ($caseChecksum = PixelChecksum($bottom[0] - 1, $bottom[1] - 1, $bottom[0] + 1, $bottom[1] + 1)) Then
				;== Boxes color check
				Sleep(100)
				$boxes = PixelSearch($box1[0], $box1[1], $box2[0], $box2[1], $iColorBox, 0)
				;== Box check
				;== Box available
				;ToolTip("Box check")
				;Sleep(100)
				If IsArray($boxes) Then
					;== Mouse clicks pick button
					MouseMove($pickButton[0], $pickButton[1], 0)

					If $casesPassed = 1 Then
						ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " case passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					Else
						ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " cases passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					EndIf

					MouseClick($MOUSE_CLICK_LEFT)
					$errorPresent = 0
					$errorTimer = 0

					;== False positive check
					While ($errorTimer <= 1000)
						While ($errorTimer <= 30)
							ToolTip(Floor($errorTimer / 10) & "% Lie Detecting")
							;== Color of pick button when it is present with the cursor over it
							$pickColor = PixelGetColor($pickButton[0] + 5, $pickButton[1] - 5)
							$errorTimer = $errorTimer + 3
							Sleep(50)
						WEnd
						While ($errorTimer <= 999)
							If (PixelGetColor($pickButton[0] + 5, $pickButton[1] - 5) = $pickColor) Then
								ToolTip(Floor($errorTimer / 10) & "% Lie Detecting")
								;== If no error message, errorPresent stays 0
								If ($errorChecksum = PixelChecksum($error[0] - 10, $error[1] - 10, $error[0] + 10, $error[1] + 10)) Then
									$errorTimer = $errorTimer + 1
								;== If error message, errorPresent counts up
								Else
									$errorPresent = $errorPresent + 1
									$errorTimer = $errorTimer + 1
								EndIf
							ElseIf (Hex(PixelGetColor($pickButton[0] + 5, $pickButton[1] - 5), 6) = "FFFFFF") Then
								$pickColor = "FF00FF"
								ToolTip(Floor($errorTimer / 10) & "% Lie Detecting")
								;== If no error message, errorPresent stays 0
								If ($errorChecksum = PixelChecksum($error[0] - 10, $error[1] - 10, $error[0] + 10, $error[1] + 10)) Then
									$errorTimer = $errorTimer + 10
								;== If error message, errorPresent counts up
								Else
									$errorPresent = $errorPresent + 1
									$errorTimer = $errorTimer + 10
								EndIf
							Else
								ToolTip(Floor($errorTimer / 10) & "% Lie Detecting")
								;== If no error message, errorPresent stays 0
								If ($errorChecksum = PixelChecksum($error[0] - 10, $error[1] - 10, $error[0] + 10, $error[1] + 10)) Then
									$errorTimer = $errorTimer + 40
								;== If error message, errorPresent counts up
								Else
									$errorPresent = $errorPresent + 1
									$errorTimer = $errorTimer + 40
								EndIf
							EndIf
							$errorTimer = $errorTimer + 10
							Sleep(50)
						WEnd
						$errorTimer = $errorTimer + 10
						Sleep(50)
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
						ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " case passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					Else
						ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " cases passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
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
								ToolTip("Pop Now Bot" & @CRLF & "Box found!" & @CRLF & "It only took " & $casesPassed + 1 & " case!")
							Else
								ToolTip("Pop Now Bot" & @CRLF & "Box found!" & @CRLF & "It only took " & $casesPassed + 1 & " cases!")
							EndIf
							FileOpen(@WorkingDir & "\Success Sound\Champions.mp3")
							SoundPlay(@WorkingDir & "\Success Sound\Champions.mp3", 1)
							FileClose(@WorkingDir & "\Success Sound\Champions.mp3")
							Sleep(50)
						WEnd
					EndIf
				;== Box unavailable
				Else
					MouseMove($nextButton[0], $nextButton[1], 0)

					If $casesPassed = 1 Then
						ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " case passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
						Sleep(50)
					Else
						ToolTip("Pop Now Bot" & @CRLF & $casesPassed & " cases passed" & @CRLF & @CRLF & "Home to pause" & @CRLF & "Escape to quit")
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
						If $caseLeftColor = PixelGetColor($box1[0], $box1[1]) Then
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
		            $hPic = GUICtrlCreatePic($sBMP_Path, 0, 0, $iX2 - $iX1 + 1, $iY2 - $iY1 + 1)
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
	; -------------
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
		If (MsgBox($MB_OKCANCEL, "Case Location", "Position your cursor on a unique color on the bottom of the case. Position will be saved in " & $iTimeout & " seconds or click cancel.",$iTimeout) = 2) Then
			Terminate()
		EndIf
		$bottom = MouseGetPos()

		GUICreate("", 200, 200) ; will create a dialog box that when displayed is centered
		$colorCase = "0x" & Hex(PixelGetColor($bottom[0], $bottom[1]), 6)
    	GUISetBkColor($colorCase)
		$textRGB = _ColorGetRGB($colorCase)
		
		$brightText = GUICtrlCreateLabel("Selected case color" & @CRLF & @CRLF & "Close this window to proceed", 10, 10)

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

		$continue = MsgBox($MB_CANCELTRYCONTINUE, "Case Location", "Position set to x = " & $bottom[0] & ", y = " & $bottom[1] & ". Continue?")
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
	;== If responce is 'Cancel', quit program
	If ($continue = 2) Then
		Terminate()
	EndIf
EndFunc ;== Setup

Func LoadPreset()
	$refreshGUI = 1
	Do
		GUICreate("Pop Now Bot Presets", 300, 360)
		Local $idPresetList = GUICtrlCreateList("", 25, 25, 250, 250)
		Local $idButton_Confirm = GUICtrlCreateButton("Confirm preset", 25, 275, 250, 25)
		Local $idButton_Delete = GUICtrlCreateButton("Delete Preset", 25, 310, 250, 25)
		;
		Local $hSearch = FileFindFirstFile("Presets\*.txt")
		;
		If $hSearch = -1 Then
    	    MsgBox($MB_SYSTEMMODAL, "", "Error: No files/directories matched the search pattern.")
    	    Return False
    	EndIf
		;
		Local $sFileName = "", $iResult = 0
		;
		While 1
    	    $sFileName = FileFindNextFile($hSearch)
    	    ; If there is no more file matching the search.
    	    If @error Then ExitLoop
			;
    	    ; Add file name to the list
    	    GUICtrlSetData($idPresetList, $sFileName)
    	WEnd
		;	
		FileClose($hSearch)
		;
		GUICtrlSetState(-1, $GUI_FOCUS)
		;
		GUISetState(@SW_SHOW)
		;
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

	;$rFileName = @WorkingDir & "\Presets\" & $lPresetName & ".txt"

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

	MsgBox($MB_SYSTEMMODAL, "Loaded Content", "Top Left of boxes [x] = " & $box1[0] & @CRLF & "Top Left of boxes [y] = " & $box1[1] & @CRLF & "Bottom Right of boxes [x] = " & $box2[0] & @CRLF & "Bottom Right of boxes [y] = " & $box2[1] & @CRLF & "Lowest point of case [x] = " & $bottom[0] & @CRLF & "Lowest point of case [y] = " & $bottom[1] & @CRLF & "Center of error banner area [x] = " & $error[0] & @CRLF & "Center of error banner area [y] = " & $error[1] & @CRLF & "Center of next button [x] = " & $next[0] & @CRLF & "Center of next button [y] = " & $next[1] & @CRLF & "Bottom quarter of next button [x] = " & $nextButton[0] & @CRLF & "Bottom quarter of next button [y] = " & $nextButton[1] & @CRLF & "Location of 'Pick One' button [x] = " & $pickButton[0] & @CRLF & "Location of 'Pick One' button [y] = " & $pickButton[1] & @CRLF & "Color of Pop Now figure = 0x" & Hex($iColorBox, 6))

	FileClose($rFileHandle)
EndFunc ;== LoadPreset

Func SavePreset()
	$sNewPreset = InputBox("Save Preset", "Enter the name of the preset.")

	; Create file
    $sFileName = @WorkingDir & "\Presets\" & $sNewPreset & ".txt"
	
	If FileExists($sFileName) Then
	    If (MsgBox($MB_YESNO, "File", "File already exists. Repalce old file?") = 6) Then
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

	MsgBox($MB_SYSTEMMODAL, "File Content", FileRead($sFileName))

    ; Close the handle returned by FileOpen.
    FileClose($sFileName)

	MsgBox($MB_OK, "Save Preset", "Preset saved." & @CRLF & "File location: " & FileGetLongName($sFileName))

EndFunc ;== SavePreset

Func Mark_Rect()

    Local $aMouse_Pos, $aMask, $aM_Mask, $iTemp
    Local $UserDLL = DllOpen("user32.dll")

; Wait until mouse button pressed
    While Not _IsPressed("01", $UserDLL)
		ToolTip("Pop Now Bot" & @CRLF & @CRLF & "Click and Drag to select area")
        Sleep(50)
    WEnd

; Get first mouse position
    $aMouse_Pos = MouseGetPos()
    $iX1 = $aMouse_Pos[0]
    $iY1 = $aMouse_Pos[1]

    Global $hRectangle_GUI = GUICreate("", @DesktopWidth, @DesktopHeight, 0, 0, $WS_POPUP, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
    GUISetBkColor(0x000000)

; Draw rectangle while mouse button pressed
    While _IsPressed("01", $UserDLL)
		ToolTip("Pop Now Bot" & @CRLF & @CRLF & "Selecting area")
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