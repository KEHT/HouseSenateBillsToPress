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

Global $sXMLinputFolderDefault = "\\alpha3\E\SOLCBILL"
Global $sTXTinputFolderDefault = "\\alpha3\E\RECSCAN"
Global $sXML2LOCexecDefault = @ScriptDir & "\ManyXML2Loc.exe "
Global $sXMLinputFolder, $sTXTinputFolder, $sXML2LOCexec

Dim $Date, $DateSelected, $ValidDate, $msg, $LocalDate, $hExtensionNumberFieldGUI, $hBlurbFileNameFieldGUI, $hAmendmentFileNameFieldGUI, $hAddButtonGUI, _
	$hCancelButtonGUI, $hBlurbAmendmentList, $hCombineButtonGUI, $hCombinePrintButtonGUI, $hBlurbFolder, $hAmendFolder, $hXML2LOCpath, $hDefault_Button, _
	$hApply_Button

fuMainGUI()

; create GUI and tabs
Func fuMainGUI()
	$hMainGUI = GUICreate("Senate " & ChrW(8212) & " Combine Files v0.9.0.0", 600, 500)
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the main GUI [X] is clicked

	$tab = GUICtrlCreateTab(5, 5, 590, 490)

	; tab 0
	$tab0 = GUICtrlCreateTabItem("Main")

	$LocalDate = _DateAdd('d', -1, _NowCalcDate())

	$Date = GUICtrlCreateMonthCal($LocalDate, 340, 50, 220, 140, $MCS_NOTODAY)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function

	$DateSelected = GUICtrlCreateLabel("Date Selected: " & $LocalDate, 340, 200, 300)

	GUISetFont(9, $FW_BOLD)
	GUICtrlCreateLabel("Extension Number", 55, 65)
	GUICtrlCreateLabel("Blurb File Name", 28, 110)
	GUICtrlCreateLabel("Amendment File Name", 175, 110)
	GUISetFont(Default, $FW_NORMAL)
	$hExtensionNumberFieldGUI = GUICtrlCreateInput("", 170, 60, 120, 22)
	$hBlurbFileNameFieldGUI = GUICtrlCreateInput("", 28, 130, 130, 22)
	$hAmendmentFileNameFieldGUI = GUICtrlCreateInput("", 175, 130, 130, 22)

	GUISetFont(Default, $FW_BOLD)
	$hAddButtonGUI = GUICtrlCreateButton("ADD", 125, 180, 80, 22)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	HotKeySet("{ENTER}", "HotKeyPressed")
	GUISetFont(Default, $FW_NORMAL)

	$hBlurbAmendmentList = GUICtrlCreateListView("", 28, 230, 544, 200, BitOR($LVS_SHOWSELALWAYS, $LVS_REPORT, $LVS_NOSORTHEADER, $LVS_EDITLABELS, $LVS_NOLABELWRAP))
	_GUICtrlListView_SetExtendedListViewStyle($hBlurbAmendmentList, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES))
	_GUICtrlListView_AddColumn($hBlurbAmendmentList, "Blurb File Name", 270, 2)
	_GUICtrlListView_AddColumn($hBlurbAmendmentList, "Amendment File Name", 270, 2)
	HotKeySet("{DELETE}", "HotKeyPressed")

	GUISetFont(Default, $FW_BOLD)
	$hCombineButtonGUI = GUICtrlCreateButton("COMBINE && SAVE", 225, 450, 150, 35)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	GUISetFont(Default, $FW_NORMAL)

	; tab 1
	$tab1 = GUICtrlCreateTabItem("Settings")

	GUICtrlCreateLabel("Default Blurb Directory", 35, 45)
	$hBlurbFolder = GUICtrlCreateInput("", 35, 65, 320, 20)
	$sXMLinputFolder = fuGetRegValsForSettings("blurb", $sXMLinputFolderDefault)
	GUICtrlSetData($hBlurbFolder, $sXMLinputFolder)

	GUICtrlCreateLabel("Default Amendment Directory", 35, 100)
	$hAmendFolder = GUICtrlCreateInput("", 35, 120, 320, 20)
	$sTXTinputFolder = fuGetRegValsForSettings("amend", $sTXTinputFolderDefault)
	GUICtrlSetData($hAmendFolder, $sTXTinputFolder)

	GUICtrlCreateLabel("Default Path to XML2LOC", 35, 155)
	$hXML2LOCpath = GUICtrlCreateInput("", 35, 175, 320, 20)
	$sXML2LOCexec = fuGetRegValsForSettings("xml2loc", $sXML2LOCexecDefault)
	GUICtrlSetData($hXML2LOCpath, $sXML2LOCexec)

	$hDefault_Button = GUICtrlCreateButton("Default", 400, 170, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$hApply_Button = GUICtrlCreateButton("Apply", 485, 170, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
;~ ========================================================================
	Local $sDriveName = "", $sDirName = "\\ALPHA3\E\SJOHNSON\SOLCBILL", $sExt = "xml"
	Local $asLocFiles[5] = [_PathMake($sDriveName, $sDirName, "alb14638", $sExt), _PathMake($sDriveName, $sDirName, "dav14d27", $sExt), _PathMake($sDriveName, $sDirName, "dav14d25", $sExt),_PathMake($sDriveName, $sDirName, "mir14414", $sExt),_PathMake($sDriveName, $sDirName, "alb14635", $sExt)]
	For $sLocFile In $asLocFiles
		GUICtrlCreateListViewItem($sLocFile & "|", $hBlurbAmendmentList)
	Next


;~ ========================================================================

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

;~ Captures certain keys activation
Func HotKeyPressed()
	Switch @HotKeyPressed ; The last hotkey pressed
		Case "{ENTER}" ; String is the {ENTER} hotkey
			If _GuiCtrlGetFocus($hMainGUI) = $hAmendmentFileNameFieldGUI Then
				GUICtrlCreateListViewItem(GUICtrlRead($hBlurbFileNameFieldGUI) & "|" & GUICtrlRead($hAmendmentFileNameFieldGUI), $hBlurbAmendmentList)
				GUICtrlSetData($hBlurbFileNameFieldGUI, "")
				GUICtrlSetData($hAmendmentFileNameFieldGUI, "")
				GUICtrlSetState($hBlurbFileNameFieldGUI, $GUI_FOCUS)
			Else
				HotKeySet("{ENTER}")
				Send("{ENTER}")
				HotKeySet("{ENTER}", "HotKeyPressed")
			EndIf
		Case "{DELETE}" ; String is the {DELETE} hotkey
			If _GuiCtrlGetFocus($hMainGUI) = $hBlurbAmendmentList Then
				_GUICtrlListView_DeleteItemsSelected($hBlurbAmendmentList)
			Else
				HotKeySet("{DELETE}")
				Send("{DELETE}")
				HotKeySet("{DELETE}", "HotKeyPressed")
			EndIf
	EndSwitch
EndFunc   ;==>HotKeyPressed

Func On_Click()
	Switch @GUI_CtrlId ; See wich item sent a message
		Case $Date
			$ValidDate = _DateIsValid(GUICtrlRead($Date))
			If $ValidDate Then
				GUICtrlSetData($DateSelected, "Date Selected: " & GUICtrlRead($Date))
			EndIf
		Case $hAddButtonGUI
			GUICtrlCreateListViewItem(GUICtrlRead($hBlurbFileNameFieldGUI) & "|" & GUICtrlRead($hAmendmentFileNameFieldGUI), $hBlurbAmendmentList)
			GUICtrlSetData($hBlurbFileNameFieldGUI, "")
			GUICtrlSetData($hAmendmentFileNameFieldGUI, "")
			GUICtrlSetState($hBlurbFileNameFieldGUI, $GUI_FOCUS)
		Case $hCombineButtonGUI
			Local $asBlurbItems = _GUICtrlListView_GetAllTextToArray($hBlurbAmendmentList)
			$extended = @extended
			If @error < 0 Then Exit MsgBox(0, 'Message', 'List empty')
			fuCreateXML2LOCstring($asBlurbItems)
		Case $hDefault_Button
			$sXMLinputFolder = $sXMLinputFolderDefault
			GUICtrlSetData($hBlurbFolder, $sXMLinputFolder)
			$sTXTinputFolder = $sTXTinputFolderDefault
			GUICtrlSetData($hAmendFolder, $sTXTinputFolder)
			$sXML2LOCexec = $sXML2LOCexecDefault
			GUICtrlSetData($hXML2LOCpath, $sXML2LOCexec)
		Case $hApply_Button
			fuApplySettingsValue($hBlurbFolder, "blurb")
			fuApplySettingsValue($hAmendFolder, "amend")
			fuApplySettingsValue($hXML2LOCpath, "xml2loc")

	EndSwitch
EndFunc   ;==>On_Click

Func _GuiCtrlGetFocus($GuiRef)
    Local $hwnd = ControlGetHandle($GuiRef, "", ControlGetFocus($GuiRef))
    Local $result = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hwnd)
    Return $result[0]
EndFunc

Func fuCreateXML2LOCstring($asXMLfiles = 0)
	Local $sXMLfileString = "", $sDrive = "", $sDir = "", $sFilename = "", $sExtension = ""
	Local $aPathSplit = 0, $asLocedFiles = 0
	If IsArray($asXMLfiles) = 0 Then
		Return
	EndIf
	For $iRow=0 To UBound($asXMLfiles, 1)-1
		$aPathSplit = _PathSplit($asXMLfiles[$iRow][0], $sDrive, $sDir, $sFilename, $sExtension)
		If StringLower($aPathSplit[4]) = ".xml" And $aPathSplit[2] <> "" Then
			$sXMLfileString &= $aPathSplit[0] & " "
			_ArrayAdd($asLocedFiles, _PathMake($aPathSplit[1], $aPathSplit[2], $aPathSplit[3], ".loc"))
		ElseIf 	StringLower($aPathSplit[4]) =".xml" And $aPathSplit[2] = "" Then
			$sXMLfileString &= _PathMake( "", $sXMLinputFolder, $aPathSplit[3], $aPathSplit[4] ) & " "
			_ArrayAdd($asLocedFiles,  _PathMake( "", $sXMLinputFolder, $aPathSplit[3], ".loc"))
		EndIf
	Next
	fuXML2LOC(StringStripWS($sXMLfileString, $STR_STRIPTRAILING))
	fuDoWok($asLocedFiles)
EndFunc

Func fuXML2LOC($sXMLfileString = "")
	If $sXMLfileString <> "" Then
		 Local $iPID = Run($sXML2LOCexec & $sXMLfileString, "")
		 Local $hWnd = WinWaitActive("ManyXML2Loc")
		 ControlClick($hWnd, "", "OK")
		 ProcessWaitClose($iPID)
	EndIf
	Return
EndFunc

Func fuDoWok($asLocFiles = 0)
	For $sLocFile In $asLocFiles
		Local $iPID = Run("F:\APPS\DoWok.exe", "", @SW_HIDE )
		Local $hWnd = WinWaitActive("DoWok")
		ControlClick($hWnd, "", "[ID:106]")
		ControlCommand($hWnd, "", "[ID:101]", "SelectString", 'B2R.wok')
		ControlSend($hWnd, "", "[ID:100]", StringLower($sLocFile), 1)
		Local $hCIfWnd = WinWaitActive("[CLASS:#32770]")
		ControlClick($hCIfWnd, "", "[CLASS:Button; INSTANCE:3]")
	;~ 		ControlSetText($hWnd, "", "[ID:100]", $sLocFile, 1)
		ControlClick($hWnd, "", "[ID:103]")
		ProcessWaitClose($iPID)
	Next
	Return
EndFunc

Func _GUICtrlListView_GetAllTextToArray($hListView)
    If Not IsHWnd($hListView) Then
        $hListView = GUICtrlGetHandle($hListView)
        If Not IsHWnd($hListView) Then Return SetError(1, 0, 0)
    EndIf
    Local $hWin, $iItems, $iSubitems
    $hWin = _WinAPI_GetParent($hListView)
    If Not $hWin Then Return SetError(2, 0, 0)
    $iItems = ControlListView($hWin, '', $hListView, 'GetItemCount')
    If Not $iItems Then Return SetError(-1, 0, 0)
    $iSubitems = ControlListView($hWin, '', $hListView, 'GetSubItemCount')
    ; If Not $iSubitems Then Return SetError(-2, 0, 0)
    If Not $iSubitems Then $iSubitems = 1
    Local $aReturn[$iItems][$iSubitems] = [[$iItems]]
    ; For $i = 0 To $iItems - 1
        ; For $j = 0 To $iSubitems - 1
            ; $aReturn[$i][$j] = ControlListView($hWin, '', $hListView, 'GetText', $i, $j)
        ; Next
    ; Next
    For $j = 0 To $iSubitems - 1
        For $i = 0 To $iItems - 1
            $aReturn[$i][$j] = ControlListView($hWin, '', $hListView, 'GetText', $i, $j)
        Next
    Next
    Return SetError(0, _WinAPI_MakeLong($iItems, $iSubitems), $aReturn)
EndFunc   ;==>_GUICtrlListView_GetAllTextToArray

; function to get input or output values from registry if they exist
Func fuGetRegValsForSettings($sFolder, $DefaultFolder)

	Local $sRegValue

	$sRegValue = RegRead("HKEY_CURRENT_USER\Software\USGPO\PED\SenateCombineFiles", $sFolder)
	If $sRegValue = "" Then
		RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\SenateCombineFiles", $sFolder, "REG_SZ", $DefaultFolder)
		Return $DefaultFolder
	Else
		Return $sRegValue
	EndIf

EndFunc   ;==>GetInputOutput

Func fuApplySettingsValue($hGUI, $sFolder)
	$cInputVal = GUICtrlRead($hGUI)
	$cInputVal = StringRegExpReplace($cInputVal, '\\* *$', '') ; strip trailing \ and spaces
	If Not FileExists($cInputVal) Then
		MsgBox(16, "Location invalid", $sFolder & " location does not exists. Enter a valid path to it.")
	Else
		If Not RegWrite("HKEY_CURRENT_USER\Software\USGPO\PED\SenateCombineFiles", $sFolder, "REG_SZ", $cInputVal) Then
			MsgBox(16, "Could not be saved", $sFolder & " location could not be saved, Error #" & @error)
		EndIf
	EndIf
	GUICtrlSetData($hGUI, $cInputVal)
EndFunc
