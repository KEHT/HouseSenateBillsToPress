;~ 12/23/2014 - sjohnson@gpo.gov - Alpha version (0.90) SenateCombineFiles to process blurbs & amendments
;~ 12/24/2014 - sjohnson@gpo.gov - Alpha version (0.901) Added Progress Bar
#include <file.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include <DateTimeConstants.au3>
#include <ProgressConstants.au3>
#include <StringConstants.au3>
#include <FontConstants.au3>
#include <GuiListView.au3>
#include <GuiEdit.au3>
#include <Constants.au3>

Opt("GUIOnEventMode", 1)

;~ Dim $yorno = 7
;~ Dim $szDrive, $szDir, $szFName, $szExt, $aFile, $sInputFileName, $sInputFile, $sInputFileText
Global $iProgress = 0

Global $sXMLinputFolderDefault = "\\alpha3\E\SOLCBILL" ; "\\alpha3\E\RiosBay\Text of Amendments\1_SOLCBILL_Run XML2LOC" ;
Global $sTXTinputFolderDefault = "\\alpha3\E\RECSCAN" ; "\\alpha3\E\RiosBay\Text of Amendments\3_RECSCAN";
Global $sXML2LOCexecDefault = @ScriptDir & "\ManyXML2Loc.exe"
Global $sDoWokexecDefault = "\\alpha3\NETAPPS\APPS\DoWok.exe" ; "\\alpha3\E\DAVE\DoWok.exe" ;
Global $sSluglineFileDefault = "\\alpha3\E\CR\NSET\slugline"
Global $sOutputDirDefault = "\\alpha3\E\CR\NSET" ; "\\alpha3\E\RiosBay\Text of Amendments\4_CR_NSET" ;
Global $sTextPadDefault = 'C:\Program Files (x86)\TextPad 4\TextPad.exe'
Global $sXMLinputFolder, $sTXTinputFolder, $sXML2LOCexec, $sDoWokexec, $sSluglineFile, $sOutputDir, $sTextPad

Dim $Date, $DateSelected, $ValidDate, $hMainGUI, $hExtensionNumberFieldGUI, $hBlurbFileNameFieldGUI, $hAmendmentFileNameFieldGUI, $hAddButtonGUI, _
		$hCancelButtonGUI, $hBlurbAmendmentList, $hCombineButtonGUI, $hCombinePrintButtonGUI, $hBlurbFolder, $hAmendFolder, $hXML2LOCpath, $hDefault_Button, _
		$hApply_Button, $hDoWokpath, $hSluglinepath, $hOutputDirpath, $hTextPad

fuMainGUI()

; create GUI and tabs
Func fuMainGUI()
	$hMainGUI = GUICreate("Senate " & ChrW(8212) & " Combine Files v0.9.0.1", 600, 500)
	GUISetOnEvent($GUI_EVENT_CLOSE, "On_Close") ; Run this function when the main GUI [X] is clicked

	$tab = GUICtrlCreateTab(5, 5, 590, 490)

	; tab 0
	$tab0 = GUICtrlCreateTabItem("Main")

;~ 	$LocalDate = _DateAdd('d', -1, _NowCalcDate())

	$Date = GUICtrlCreateMonthCal(_NowCalcDate(), 340, 50, 220, 140, $MCS_NOTODAY)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function

	$DateSelected = GUICtrlCreateLabel("Date Selected: " & _NowCalcDate(), 340, 200, 300)

	GUISetFont(9, $FW_BOLD)
	GUICtrlCreateLabel("Extension Number", 55, 65)
	GUICtrlCreateLabel("Blurb File Name", 28, 110)
	GUICtrlCreateLabel("Amendment File Name", 175, 110)
	GUISetFont(Default, $FW_NORMAL)
	$hExtensionNumberFieldGUI = GUICtrlCreateInput("", 170, 60, 35, 22)
	_GUICtrlEdit_SetLimitText($hExtensionNumberFieldGUI, 3)
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

	GUICtrlCreateLabel("Default Path to DoWok", 35, 210)
	$hDoWokpath = GUICtrlCreateInput("", 35, 230, 320, 20)
	$sDoWokexec = fuGetRegValsForSettings("dowok", $sDoWokexecDefault)
	GUICtrlSetData($hDoWokpath, $sDoWokexec)

	GUICtrlCreateLabel("Default Path to Slugline", 35, 265)
	$hSluglinepath = GUICtrlCreateInput("", 35, 285, 320, 20)
	$sSluglineFile = fuGetRegValsForSettings("slugline", $sSluglineFileDefault)
	GUICtrlSetData($hSluglinepath, $sSluglineFile)

	GUICtrlCreateLabel("Default Output Path", 35, 320)
	$hOutputDirpath = GUICtrlCreateInput("", 35, 340, 320, 20)
	$sOutputDir = fuGetRegValsForSettings("output", $sOutputDirDefault)
	GUICtrlSetData($hOutputDirpath, $sOutputDir)

	GUICtrlCreateLabel("Default Path to TextPad", 35, 375)
	$hTextPad = GUICtrlCreateInput("", 35, 395, 320, 20)
	$sTextPad = fuGetRegValsForSettings("textpad", $sTextPadDefault)
	GUICtrlSetData($hTextPad, $sTextPad)

	$hDefault_Button = GUICtrlCreateButton("Default", 400, 225, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
	$hApply_Button = GUICtrlCreateButton("Apply", 485, 225, 75)
	GUICtrlSetOnEvent(-1, "On_Click") ; Call a common button function
;~ ========================================================================
;~ 	Local $sDriveName = "", $sDirName = "\\ALPHA3\E\SJOHNSON\SOLCBILL", $sBlurbDirName = "\\ALPHA3\E\SJOHNSON\RECSCAN", $sExt = "xml", $sBlurbExt = "txt"
;~ 	Local $asLocFiles[4] = [_PathMake($sDriveName, $sDirName, "alb14638", $sExt), _PathMake($sDriveName, $sDirName, "dav14d27", $sExt), _PathMake($sDriveName, $sDirName, "dav14d25", $sExt),_PathMake($sDriveName, $sDirName, "mir14414", $sExt)]
;~ 	Local $asBlurbFiles[4] = [_PathMake($sDriveName, $sBlurbDirName, "SA3783", $sBlurbExt), _PathMake($sDriveName, $sBlurbDirName, "SA3784", $sBlurbExt), _PathMake($sDriveName, $sBlurbDirName, "SA3785", $sBlurbExt), _PathMake($sDriveName, $sBlurbDirName, "SA3786", $sBlurbExt) ]
;~ 	For $iI = 0 to 3 Step 1
;~ 		GUICtrlCreateListViewItem($asBlurbFiles[$iI] & "|" & $asLocFiles[$iI], $hBlurbAmendmentList)
;~ 	Next

;~ ========================================================================

	GUISetState()

	; Run the GUI until the dialog is closed
	While 1
		Sleep(10)
	WEnd
EndFunc   ;==>fuMainGUI

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
				fuPostFiles($hBlurbAmendmentList, GUICtrlRead($hBlurbFileNameFieldGUI), GUICtrlRead($hAmendmentFileNameFieldGUI))
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
			fuPostFiles($hBlurbAmendmentList, GUICtrlRead($hBlurbFileNameFieldGUI), GUICtrlRead($hAmendmentFileNameFieldGUI))
			GUICtrlSetData($hBlurbFileNameFieldGUI, "")
			GUICtrlSetData($hAmendmentFileNameFieldGUI, "")
			GUICtrlSetState($hBlurbFileNameFieldGUI, $GUI_FOCUS)
		Case $hCombineButtonGUI
			ProgressOn("Task Progress", "Processing Blurbs && Amendments", "Combining...", Default, Default, $DLG_NOTITLE)
			Local $iExtensionNumber = Int(GUICtrlRead($hExtensionNumberFieldGUI))
			If $iExtensionNumber < 1 Then Return MsgBox($MB_ICONERROR, 'Error', 'Extension Number is not between 001 and 999 !!!')
			Local $sExtensionNumber = StringFormat("%03i", $iExtensionNumber)
			Local $asBlurbItems = _GUICtrlListView_GetAllTextToArray($hBlurbAmendmentList)
			$extended = @extended
			If @error < 0 Then Return MsgBox($MB_ICONERROR, 'Error', 'File List is  Empty !!!' & @CRLF & "@error = " & @error & ", @extended = " & @extended)
			Local $tLocedFiles = fuCreateXML2LOCstring($asBlurbItems)
			fuDoWok($tLocedFiles)
			fuCombineFiles($tLocedFiles, $sExtensionNumber)
		Case $hDefault_Button
			$sXMLinputFolder = $sXMLinputFolderDefault
			GUICtrlSetData($hBlurbFolder, $sXMLinputFolder)
			$sTXTinputFolder = $sTXTinputFolderDefault
			GUICtrlSetData($hAmendFolder, $sTXTinputFolder)
			$sXML2LOCexec = $sXML2LOCexecDefault
			GUICtrlSetData($hXML2LOCpath, $sXML2LOCexec)
			$sDoWokexec = $sDoWokexecDefault
			GUICtrlSetData($hDoWokpath, $sDoWokexec)
			$sSluglineFile = $sSluglineFileDefault
			GUICtrlSetData($hSluglinepath, $sSluglineFile)
			$sOutputDir = $sOutputDirDefault
			GUICtrlSetData($hOutputDirpath, $sOutputDir)
			$sTextPad = $sTextPadDefault
			GUICtrlSetData($hTextPad, $sTextPad)
			ContinueCase
		Case $hApply_Button
			fuApplySettingsValue($hBlurbFolder, "blurb")
			fuApplySettingsValue($hAmendFolder, "amend")
			fuApplySettingsValue($hXML2LOCpath, "xml2loc")
			fuApplySettingsValue($hDoWokpath, "dowok")
			fuApplySettingsValue($hSluglinepath, "slugline")
			fuApplySettingsValue($hOutputDirpath, "output")
			fuApplySettingsValue($hTextPad, "textpad")
	EndSwitch
EndFunc   ;==>On_Click

Func _GuiCtrlGetFocus($GuiRef)
	Local $hwnd = ControlGetHandle($GuiRef, "", ControlGetFocus($GuiRef))
	Local $result = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hwnd)
	Return $result[0]
EndFunc   ;==>_GuiCtrlGetFocus

Func fuCreateXML2LOCstring($asInputFiles = 0)
	Local $sXMLfileString = "", $sDrive = "", $sDir = "", $sFilename = "", $sExtension = ""
	Local $aPathSplit = 0 ;, $asLocedFiles = [0]
	Local $tLocedFiles = ObjCreate('Scripting.Dictionary')
	If IsArray($asInputFiles) = 0 Then
		Return
	Else
		Local $iFileCount = UBound($asInputFiles, 1) - 1
	EndIf
	For $iRow = 0 To $iFileCount
		$aPathSplit = _PathSplit($asInputFiles[$iRow][1], $sDrive, $sDir, $sFilename, $sExtension)
		If StringLower($sExtension) = ".xml" Then
			$sXMLfileString &= $aPathSplit[0] & " "
			$tLocedFiles.Add($asInputFiles[$iRow][0], _PathMake($sDrive, $sDir, $sFilename, ".loc"))
		Else
			$tLocedFiles.Add($asInputFiles[$iRow][0], _PathMake($sDrive, $sDir, $sFilename, $sExtension))
		EndIf
		$iProgress += (33 / $iFileCount)
		ProgressSet($iProgress)
	Next
	fuXML2LOC(StringStripWS($sXMLfileString, $STR_STRIPTRAILING))
	Return $tLocedFiles
EndFunc   ;==>fuCreateXML2LOCstring

Func fuXML2LOC($sXMLfileString = "")
	If $sXMLfileString <> "" Then
		Local $iPID = Run($sXML2LOCexec & " " & $sXMLfileString, "")
		Local $hwnd = WinWaitActive("ManyXML2Loc")
		ControlClick($hwnd, "", "[ID:2]")
		$hwnd = WinWaitActive("ManyXML2Loc", "", 1)
		ControlClick($hwnd, "", "[ID:2]")
		ProcessWaitClose($iPID)
	EndIf
	Return
EndFunc   ;==>fuXML2LOC

Func fuDoWok($tLocFiles = 0)
	Local $sDrive = "", $sDir = "", $sFilename = "", $sExtension = "", $sDoWokString = '"'
	Local $iKeyCount = 0
	For $sBlrbKey In $tLocFiles.Keys
		_PathSplit($tLocFiles($sBlrbKey), $sDrive, $sDir, $sFilename, $sExtension)
		If $sExtension = ".loc" Then
			$sDoWokString &= StringLower($tLocFiles($sBlrbKey)) & @TAB
		EndIf
		$iProgress += (33 / $tLocFiles.Count)
		ProgressSet($iProgress)
	Next
;~ 	StringStripWS($sDoWokString, BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING))
	$sDoWokString = StringTrimRight($sDoWokString, 1)
	$sDoWokString &= '"' & ' B2R.wok -O'
;~ 	MsgBox ( $MB_SYSTEMMODAL, "DoWok String", $sDoWokString )
	Local $iPID = Run($sDoWokexec & " " & $sDoWokString)
	ProcessWaitClose($iPID)
	Return
EndFunc   ;==>fuDoWok

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

	If Not $iSubitems Then $iSubitems = 1
	Local $aReturn[$iItems][$iSubitems] = [[$iItems]]

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

EndFunc   ;==>fuGetRegValsForSettings

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
	Return
EndFunc   ;==>fuApplySettingsValue

Func fuPostFiles($hGUI, $sBlurb = "", $sAmend = "")
	Local $sDrive = "", $sDir = "", $sFilename = "", $sExtension = ""

	_PathSplit($sBlurb, $sDrive, $sDir, $sFilename, $sExtension)
	If $sDir = "" Then
		$sDir = $sTXTinputFolder
	EndIf
	If $sExtension = "" Then
		$sExtension = ".txt"
	EndIf
	$sBlurb = _PathMake($sDrive, $sDir, $sFilename, $sExtension)

	Dim $sDrive = "", $sDir = "", $sFilename = "", $sExtension = ""
	_PathSplit($sAmend, $sDrive, $sDir, $sFilename, $sExtension)
	If $sDir = "" Then
		$sDir = $sXMLinputFolder
	EndIf
	If $sExtension = "" Then
		$sExtension = ".loc"
	EndIf
	$sAmend = _PathMake($sDrive, $sDir, $sFilename, $sExtension)

	GUICtrlCreateListViewItem($sBlurb & "|" & $sAmend, $hGUI)
	Return
EndFunc   ;==>fuPostFiles

Func fuCombineFiles($tFileList = 0, $sExtensionNumber = '')
	Local $sComboText = ''

	Local $aMonths[13] = ["00", "JA", "FE", "MR", "AP", "MY", "JN", "JY", "AU", "SE", "OC", "NO", "DE"]
	Local $cDay = GUICtrlRead($Date)
	Local $nMonth = Number(StringRegExpReplace($cDay, '(\d{4})/(\d{2})/(\d{2})', '$2'))
	Local $cTempDay = StringRegExpReplace($cDay, '(\d{4})/(\d{2})/(\d{2})', '$3')
	Local $sOutputFileName = "A" & $cTempDay & $aMonths[$nMonth] & "6"
	Local $iKeyCount = 0

	Local $hFileOpen = FileOpen($sSluglineFile, BitOR($FO_READ, $FO_UTF8_FULL))
	If $hFileOpen < 0 Then Return MsgBox($MB_ICONERROR, 'Error', 'Could Not Open Slugline File!!!')
	$sComboText &= FileRead($hFileOpen)
	FileClose($hFileOpen)

	For $sBlrbKey In $tFileList.Keys
		$sComboText &= @CRLF & @CRLF & @CRLF & @CRLF
		$hFileOpen = FileOpen(StringLower($sBlrbKey), BitOR($FO_READ, $FO_UTF8_FULL))
		If $hFileOpen < 0 Then Return MsgBox($MB_ICONERROR, 'Error', 'Could Not Open ' & $sBlrbKey & ' File!!!')
		$sComboText &= StringRegExpReplace(FileRead($hFileOpen), "(?sm)Mr.{1,2}\s\b([A-Z]+[a-z]+[a-zA-Z]*)\b\ssubmitted", "T5$1K")
		FileClose($hFileOpen)

		$sComboText &= @CRLF
		$hFileOpen = FileOpen(StringLower($tFileList($sBlrbKey)), BitOR($FO_READ, $FO_UTF8_FULL))
		If $hFileOpen < 0 Then Return MsgBox($MB_ICONERROR, 'Error', 'Could Not Open ' & $tFileList($sBlrbKey) & ' File!!!')
		$sComboText &= FileRead($hFileOpen)
		FileClose($hFileOpen)
		$iProgress += (33 / $tFileList.Keys)
		ProgressSet($iProgress)
	Next
	$sComboText = StringRegExpReplace($sComboText, "(?sm)(S7781.*?S0634)", "") ; Remove text between two bell codes
	$sComboText = StringRegExpReplace($sComboText, "(?sm)(S0634\s*S0634)", "S0634") ; Remove double S0634 bell code

	; Output of combined files to a single long file for debugging purposes
	Local $sOutputFile = $sOutputDir & '\' & $sOutputFileName & '.' & $sExtensionNumber
	If FileExists($sOutputFile) Then Return MsgBox($MB_ICONERROR, 'Error', 'File:' & @CRLF & $sOutputFile & @CRLF & 'already exists!!!')
	If FileWrite($sOutputFile, $sComboText) = 0 Then Return MsgBox($MB_ICONERROR, 'Error', 'File could not be saved:' & @CRLF & $sOutputFile)

;~ 	MsgBox(0, "Textpad String", $sTextPad)
	Run($sTextPad & " " & $sOutputFile, @WindowsDir, @SW_SHOWDEFAULT)
	ProgressSet(100, "Done!")
	$iProgress = 0
	Sleep(750)
	ProgressOff()
	Return
EndFunc   ;==>fuCombineFiles
