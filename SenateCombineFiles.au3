#include <file.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include <DateTimeConstants.au3>
#include <ProgressConstants.au3>
#include <StringConstants.au3>
#include <objDictonary.au3>

Opt("GUIOnEventMode", 1)

Dim $yorno = 7
Dim $szDrive, $szDir, $szFName, $szExt, $aFile, $sInputFileName, $sInputFile, $sInputFileText
Global $hMainGUI

Global $sXMLinputFolderDefault = "\\alpha3\E\SOLCBIL"
Global $sTXTinputFolderDefault = "\\alpha3\E\RECSCAN"
Global $sInputFolder, $sOutputFolder

Dim $Date, $DateSelected, $ValidDate, $msg, $LocalDate

fuMainGUI()

; create GUI and tabs
Func fuMainGUI()
	$hMainGUI = GUICreate("Senate — Combine Files v0.9.0.0", 600, 400)
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the main GUI [X] is clicked

	$LocalDate = _DateAdd('d', -1, _NowCalcDate())
	$Date = GUICtrlCreateMonthCal($LocalDate, 365, 17, 220, 140, $MCS_NOTODAY)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function

	GUISetState()

	; Run the GUI until the dialog is closed
	While 1
		Sleep(10)
	WEnd
EndFunc

Func On_Close()
	Switch @GUI_WinHandle ; See which GUI sent the CLOSE message
		Case $hMainGUI
			Exit ; If it was this GUI - we exit <<<<<<<<<<<<<<<

	EndSwitch
EndFunc   ;==>On_Close

Func On_Click()
	Switch @GUI_CtrlId ; See wich item sent a message
		Case $Date
			$ValidDate = _DateIsValid(GUICtrlRead($Date))
			If $ValidDate Then
				GUICtrlSetData($DateSelected, "Date Selected: " & GUICtrlRead($Date))
			EndIf


	EndSwitch
EndFunc   ;==>On_Click