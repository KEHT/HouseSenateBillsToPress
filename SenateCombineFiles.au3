#include <file.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include <DateTimeConstants.au3>
#include <ProgressConstants.au3>
#include <StringConstants.au3>
#include <objDictonary.au3>
#include <FontConstants.au3>

Opt("GUIOnEventMode", 1)

Dim $yorno = 7
Dim $szDrive, $szDir, $szFName, $szExt, $aFile, $sInputFileName, $sInputFile, $sInputFileText
Global $hMainGUI

Global $sXMLinputFolderDefault = "\\alpha3\E\SOLCBIL"
Global $sTXTinputFolderDefault = "\\alpha3\E\RECSCAN"
Global $sInputFolder, $sOutputFolder

Dim $Date, $DateSelected, $ValidDate, $msg, $LocalDate, $hExtensionNumberFieldGUI, $hExtensionLabelGUI, $hBlurbFileNameLabelGUI, $hAmendmentFileNameNGUI, _
	$hBlurbFileNameFieldGUI, $hAmendmentFileNameFieldGUI, $hAddButtonGUI, $hCancelButtonGUI, $hBlurbAmendmentList, $hCombineButtonGUI, $hCombinePrintButtonGUI

fuMainGUI()

; create GUI and tabs
Func fuMainGUI()
	$hMainGUI = GUICreate("Senate — Combine Files v0.9.0.0", 600, 500)
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the main GUI [X] is clicked

	GUICtrlCreateLabel("Choose Date Below:", 362, 7, 300)
	$LocalDate = _DateAdd('d', -1, _NowCalcDate())

	$Date = GUICtrlCreateMonthCal($LocalDate, 362, 25, 220, 140, $MCS_NOTODAY)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function

	$DateSelected = GUICtrlCreateLabel("Date Selected: " & $LocalDate, 362, 175, 300)

	GUISetFont(9, $FW_BOLD)
	$hExtensionLabelGUI = GUICtrlCreateLabel("Extension Number", 55, 30)
	$hBlurbFileNameLabelGUI = GUICtrlCreateLabel("Blurb File Name", 20, 80)
	$hAmendmentFileNameNGUI = GUICtrlCreateLabel("Amendment File Name", 175, 80)
	GUISetFont(Default, $FW_NORMAL)
	$hExtensionNumberFieldGUI = GUICtrlCreateInput("", 170, 25, 120, 22)
	$hBlurbFileNameFieldGUI = GUICtrlCreateInput("", 20, 100, 130, 22)
	$hAmendmentFileNameFieldGUI = GUICtrlCreateInput("", 175, 100, 130, 22)

	GUISetFont(Default, $FW_BOLD)
	$hAddButtonGUI = GUICtrlCreateButton("ADD", 75, 155, 80, 22)
	$hCancelButtonGUI = GUICtrlCreateButton("CANCEL", 175, 155, 80, 22)
	GUISetFont(Default, $FW_NORMAL)

	$hBlurbAmendmentList = GUICtrlCreateListView("", 20, 210, 562, 220, BitOR($LVS_SHOWSELALWAYS, $LVS_REPORT, $LVS_NOSORTHEADER, $LVS_EDITLABELS, $LVS_NOLABELWRAP))
	_GUICtrlListView_SetExtendedListViewStyle($hBlurbAmendmentList, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
	_GUICtrlListView_AddColumn($hBlurbAmendmentList, "Blurb File Name", 270, 2)
	_GUICtrlListView_AddColumn($hBlurbAmendmentList, "Amendment File Name", 270, 2)

	GUISetFont(Default, $FW_BOLD)
	$hCombineButtonGUI = GUICtrlCreateButton("COMBINE", 200, 450, 80, 35)
	$hCombinePrintButtonGUI = GUICtrlCreateButton("COMBINE && PRINT", 300, 450, 150, 35)
	GUISetFont(Default, $FW_NORMAL)

	; Create array and fill Left listview
	Global $aLV_List_Left[20]
	For $i = 0 To UBound($aLV_List_Left) - 1
		If Mod($i, 5) Then
			$aLV_List_Left[$i] = "Tomsfgssssssssss " & $i & "|Dicksfffffffffffff " & $i
		Else
			$aLV_List_Left[$i] = "Tom " & $i & "|Harry " & $i
		EndIf
		GUICtrlCreateListViewItem($aLV_List_Left[$i], $hBlurbAmendmentList)
	Next

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