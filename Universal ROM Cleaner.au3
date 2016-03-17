#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Ressources\Universal_Rom_Cleaner.ico
#AutoIt3Wrapper_Outfile=..\BIN\Universal_Rom_Cleaner.exe
#AutoIt3Wrapper_Outfile_x64=..\BIN\Universal_Rom_Cleaner64.exe
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=Nettoyeur de Rom Universel
#AutoIt3Wrapper_Res_Fileversion=1.0.0.1
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_LegalCopyright=LEGRAS David
#AutoIt3Wrapper_Res_Language=1036
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=y
#AutoIt3Wrapper_Run_Tidy=y
#Tidy_Parameters=/reel
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;*************************************************************************
;**									**
;**						Universal Rom Cleaner	**
;**						LEGRAS David		**
;**									**
;*************************************************************************

;Définition des librairies
;-------------------------
#include <String.au3>
#include <Array.au3>
#include <File.au3>
#include <GuiListView.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <ComboConstants.au3>

#include "./Include/_MultiLang.au3"
#include "./Include/_GUIListViewEx.au3"
#include "./Include/_ExtMsgBox.au3"
#include "./Include/_Trim.au3"

;FileInstall
;-----------
Global $SOURCE_DIRECTORY = @ScriptDir
If Not _FileCreate($SOURCE_DIRECTORY & "\test") Then ;Vérification des droits en écriture
	$SOURCE_DIRECTORY = @AppDataDir & "\Universal_ROM_Tools"
	DirCreate($SOURCE_DIRECTORY)
Else
	FileDelete($SOURCE_DIRECTORY & "\test")
EndIf
DirCreate($SOURCE_DIRECTORY & "\LanguageFiles")
DirCreate($SOURCE_DIRECTORY & "\Ressources")
FileInstall(".\URC-config.ini", $SOURCE_DIRECTORY & "\URC-config.ini")
FileInstall(".\LanguageFiles\URC-ENGLISH.XML", $SOURCE_DIRECTORY & "\LanguageFiles\URC-ENGLISH.XML")
FileInstall(".\LanguageFiles\URC-FRENCH.XML", $SOURCE_DIRECTORY & "\LanguageFiles\URC-FRENCH.XML")
FileInstall(".\Ressources\Universal_Rom_Cleaner.ico", $SOURCE_DIRECTORY & "\Ressources\Universal_Rom_Cleaner.ico")

;Définition des Variables
;-------------------------
Global $PathConfigINI = $SOURCE_DIRECTORY & "\URC-config.ini"
Global $LANG_DIR = $SOURCE_DIRECTORY & "\LanguageFiles"; Where we are storing the language files.
Global $user_lang = IniRead($PathConfigINI, "LAST_USE", "$user_lang", "default")
Global $path_LOG = IniRead($PathConfigINI, "GENERAL", "Path_LOG", $SOURCE_DIRECTORY & "\log.txt")
Global $path_SIMUL = IniRead($PathConfigINI, "GENERAL", "path_SIMUL", $SOURCE_DIRECTORY & "\simulation.txt")
Local $V_ROMPath, $I_LV_ATTRIBUTE, $I_LV_SUPPRESS, $I_LV_IGNORE, $A_ROMList, $A_ROMAttribut
If @Compiled Then
	$Rev = "BETA " & FileGetVersion(@ScriptFullPath)
Else
	$Rev = 'In Progress'
EndIf

;---------;
;Principal;
;---------;

_LANG_LOAD($LANG_DIR, $user_lang) ;Chargement de la langue par défaut

#Region ### START Koda GUI section ### Form= ;Création de l'interface
$F_UniversalCleaner = GUICreate(_MultiLang_GetText("main_gui"), 523, 343, 192, 124)
GUISetBkColor(0x34495C)
$H_MF = GUICtrlCreateMenu(_MultiLang_GetText("mnu_file"))
$H_MF_ROM = GUICtrlCreateMenuItem(_MultiLang_GetText("mnu_file_roms"), $H_MF)
$H_MF_LANGUE = GUICtrlCreateMenuItem(_MultiLang_GetText("mnu_file_langue"), $H_MF)
$H_MF_Separation = GUICtrlCreateMenuItem("", $H_MF)
$H_MF_Exit = GUICtrlCreateMenuItem(_MultiLang_GetText("mnu_file_exit"), $H_MF)
$H_MA = GUICtrlCreateMenu(_MultiLang_GetText("mnu_action"))
$H_MA_SIMULATION = GUICtrlCreateMenuItem(_MultiLang_GetText("mnu_action_simulation"), $H_MA)
$H_MA_CLEAN = GUICtrlCreateMenuItem(_MultiLang_GetText("mnu_action_clean"), $H_MA)
$H_MH = GUICtrlCreateMenu(_MultiLang_GetText("mnu_help"))
$H_MH_About = GUICtrlCreateMenuItem(_MultiLang_GetText("mnu_help_about"), $H_MH)
$H_LV_ATTRIBUTE = GUICtrlCreateListView(_MultiLang_GetText("lv_attribute"), 8, 8, 250, 302, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetExtendedListViewStyle($H_LV_ATTRIBUTE, $LVS_EX_FULLROWSELECT)
_GUICtrlListView_SetColumnWidth($H_LV_ATTRIBUTE, 0, 225)
$H_LV_SUPPRESS = GUICtrlCreateListView(_MultiLang_GetText("lv_suppress"), 264, 8, 250, 150, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetExtendedListViewStyle($H_LV_SUPPRESS, $LVS_EX_FULLROWSELECT)
_GUICtrlListView_SetColumnWidth($H_LV_SUPPRESS, 0, 225)
$H_LV_IGNORE = GUICtrlCreateListView(_MultiLang_GetText("lv_ignore"), 264, 160, 250, 150, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetExtendedListViewStyle($H_LV_IGNORE, $LVS_EX_FULLROWSELECT)
_GUICtrlListView_SetColumnWidth($H_LV_IGNORE, 0, 225)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1 ; Gestion de l'interface
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $H_MF_ROM ;Menu Fichier/Charger le repertoire des ROMs
			_GUIListViewEx_Close(0)
			_GUICtrlListView_DeleteAllItems($H_LV_ATTRIBUTE)
			_GUICtrlListView_DeleteAllItems($H_LV_SUPPRESS)
			_GUICtrlListView_DeleteAllItems($H_LV_IGNORE)
			$V_ROMPath = FileSelectFolder(_MultiLang_GetText("win_sel_rom_Title"), "", $FSF_CREATEBUTTON, "C:\")
			If StringRight($V_ROMPath, 1) <> '\' Then $V_ROMPath = $V_ROMPath & '\'
			$A_ROMList = _CREATEARRAY_ROM($V_ROMPath)
			$A_ROMAttribut = _CREATEARRAY_ATTRIBUT($A_ROMList)
			For $B_ROMAttribut = 0 To UBound($A_ROMAttribut) - 1
				_GUICtrlListView_AddItem($H_LV_ATTRIBUTE, $A_ROMAttribut[$B_ROMAttribut])
			Next
			$I_LV_ATTRIBUTE = _GUIListViewEx_Init($H_LV_ATTRIBUTE, $A_ROMAttribut, 0, 0, True)
			$I_LV_SUPPRESS = _GUIListViewEx_Init($H_LV_SUPPRESS, "", 0, 0, True)
			$I_LV_IGNORE = _GUIListViewEx_Init($H_LV_IGNORE, "", 0, 0, True)
			_GUIListViewEx_MsgRegister() ;Register pour le drag&drop
			_GUIListViewEx_SetActive(1) ;Activation de la LV de gauche
		Case $H_MF_LANGUE ;Menu Fichier/Langues
			_LANG_LOAD($LANG_DIR, -1)
			_GUI_REFRESH()
		Case $GUI_EVENT_CLOSE, $H_MF_Exit ;Quitter
			Exit
		Case $H_MA_SIMULATION ;Menu Action/Simulation
			$A_ROMList = _MOVE_ROM($V_ROMPath, $I_LV_ATTRIBUTE, $A_ROMList)
			$A_ROMList = _SUPPR_ROM($V_ROMPath, $I_LV_SUPPRESS, $A_ROMList)
			_CLEAN_ROM($V_ROMPath, $A_ROMList, 0)
		Case $H_MA_CLEAN ;Menu Action/Nettoyage
			$A_ROMList = _MOVE_ROM($V_ROMPath, $I_LV_ATTRIBUTE, $A_ROMList)
			$A_ROMList = _SUPPR_ROM($V_ROMPath, $I_LV_SUPPRESS, $A_ROMList)
			_CLEAN_ROM($V_ROMPath, $A_ROMList, 1)
		Case $H_MH_About ;Menu Aide/A propos
			$sMsg = "UNIVERSAL ROM CLEANER - " & $Rev & @CRLF
			$sMsg &= _MultiLang_GetText("win_About_By") & @CRLF & @CRLF
			$sMsg &= _MultiLang_GetText("win_About_Thanks") & @CRLF
			$sMsg &= "http://www.screenzone.fr/" & @CRLF
			$sMsg &= "http://www.screenscraper.fr/" & @CRLF
			$sMsg &= "http://www.recalbox.com/" & @CRLF
			$sMsg &= "http://www.emulationstation.org/" & @CRLF
			_ExtMsgBoxSet(1, 2, 0x34495c, 0xFFFF00, 10, "Arial")
			_ExtMsgBox($EMB_ICONINFO, "OK", _MultiLang_GetText("win_About_Title"), $sMsg, 15)
	EndSwitch
WEnd

;---------;
;Fonctions;
;---------;

Func _GUI_REFRESH() ;Rafraichissement de l'interface
	GUICtrlSetData($H_MF, _MultiLang_GetText("mnu_file"))
	GUICtrlSetData($H_MF_ROM, _MultiLang_GetText("mnu_file_roms"))
	GUICtrlSetData($H_MF_LANGUE, _MultiLang_GetText("mnu_file_langue"))
	GUICtrlSetData($H_MF_Exit, _MultiLang_GetText("mnu_file_exit"))
	GUICtrlSetData($H_MA, _MultiLang_GetText("mnu_action"))
	GUICtrlSetData($H_MA_SIMULATION, _MultiLang_GetText("mnu_action_simulation"))
	GUICtrlSetData($H_MA_CLEAN, _MultiLang_GetText("mnu_action_clean"))
	GUICtrlSetData($H_MH, _MultiLang_GetText("mnu_help"))
	GUICtrlSetData($H_MH_About, _MultiLang_GetText("mnu_help_about"))
	GUICtrlSetData($H_LV_ATTRIBUTE, _MultiLang_GetText("lv_attribute"))
	GUICtrlSetData($H_LV_SUPPRESS, _MultiLang_GetText("lv_suppress"))
	GUICtrlSetData($H_LV_IGNORE, _MultiLang_GetText("lv_ignore"))
EndFunc   ;==>_GUI_REFRESH

Func _CREATEARRAY_ROM($V_ROMPath) ;Création de la liste des ROMs (Chemin des ROMs)
	Local $A_ROMList = _FileListToArray($V_ROMPath, "*.*z*")
	ProgressOn(_MultiLang_GetText("prbr_createarray_rom_title"), "", "0%")
	If @error = 1 Then
		MsgBox($MB_SYSTEMMODAL, "", $V_ROMPath & " - Path was invalid.")
		Exit
	EndIf
	If @error = 4 Then
		MsgBox($MB_SYSTEMMODAL, "", "No file(s) were found.")
		Exit
	EndIf
	For $B_COLINSRT = 1 To 4
		_ArrayColInsert($A_ROMList, $B_COLINSRT)
	Next
	For $B_ROMList = 0 To UBound($A_ROMList) - 1
		$V_ProgressPRC = Round(($B_ROMList * 100) / (UBound($A_ROMList) - 1))
		ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_createarray_rom_progress") & $V_ProgressPRC & "%")
		$A_ROMList[$B_ROMList][3] = $A_ROMList[$B_ROMList][0]
		$A_ROMList[$B_ROMList][0] = StringReplace($A_ROMList[$B_ROMList][0], '[', '(')
		$A_ROMList[$B_ROMList][0] = StringReplace($A_ROMList[$B_ROMList][0], ']', ')')
		$TMP_Path = StringSplit($A_ROMList[$B_ROMList][0], "(")
		If $TMP_Path[1] = "" Then
			$TMP_Path = StringSplit($A_ROMList[$B_ROMList][0], ")")
			$TMP_Path = StringSplit($TMP_Path[2], "(")
		EndIf
		$A_ROMList[$B_ROMList][1] = _ALLTRIM($TMP_Path[1])
	Next
;~ 	_ArrayDisplay($A_ROMList, '$A_ROMList Full') ; Debug
	_ArrayDelete($A_ROMList, "0")
	_ArraySort($A_ROMList)
;~ 	_ArrayDisplay($A_ROMList, '$A_ROMList Clean & Sorted') ; Debug
	ProgressOff()
	Return $A_ROMList
EndFunc   ;==>_CREATEARRAY_ROM

Func _CREATEARRAY_ATTRIBUT($A_ROMList) ;Création de la liste des Attributs (Array des ROMs)
	Local $A_ROMAttribut[1]
	ProgressOn(_MultiLang_GetText("prbr_createarray_attribut_title"), "", "0%")
	For $B_ROMList = 0 To UBound($A_ROMList) - 1
		$V_ProgressPRC = Round(($B_ROMList * 100) / (UBound($A_ROMList) - 1))
		ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_createarray_attribut_progress") & $V_ProgressPRC & "%")
		_ArrayAdd($A_ROMAttribut, _ArrayToString(_StringBetween($A_ROMList[$B_ROMList][0], "(", ")")))
	Next
;~ 	_ArrayDisplay($A_ROMAttribut, '$A_ROMAttribut Full') ; Debug
	$A_ROMAttribut = _ArrayUnique($A_ROMAttribut)
;~ 	_ArrayDisplay($A_ROMAttribut, '$A_ROMAttribut Unique') ; Debug
	_ArrayDelete($A_ROMAttribut, "0;1")
	_ArraySort($A_ROMAttribut)
;~ 	_ArrayDisplay($A_ROMAttribut, '$A_ROMAttribut Clean & Sorted') ; Debug
	ProgressOff()
	Return $A_ROMAttribut
EndFunc   ;==>_CREATEARRAY_ATTRIBUT

Func _MOVE_ROM($V_ROMPath, $I_LV_ATTRIBUTE, $A_ROMList) ;Définition des ROMs à deplacer (Chemin des ROMs, Indexe de la LV des attributs triés, Array des ROMs)
	Local $A_TEMP_RomList
	ProgressOn(_MultiLang_GetText("prbr_move_rom_title"), "", "0%")
	$A_LV_ATTRIBUTE = _GUIListViewEx_ReturnArray($I_LV_ATTRIBUTE)
	_ArrayReverse($A_LV_ATTRIBUTE)
	For $B_ROMList = 0 To UBound($A_ROMList) - 1
		$A_ROMList[$B_ROMList][2] = 0
	Next
;~ 	_ArrayDisplay($A_ROMList, '$A_ROMList Reversed & Completed') ; Debug
	For $B_ROMList = 0 To UBound($A_ROMList) - 1
		$V_ProgressPRC = Round(($B_ROMList * 100) / (UBound($A_ROMList) - 1))
		ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_move_rom_progress") & $V_ProgressPRC & "%")
		For $B_LV_ATTRIBUTE = 0 To UBound($A_LV_ATTRIBUTE) - 1
			If StringInStr($A_ROMList[$B_ROMList][0], "(" & $A_LV_ATTRIBUTE[$B_LV_ATTRIBUTE] & ")") > 0 Then
				$A_ROMList[$B_ROMList][2] = $A_ROMList[$B_ROMList][2] + (($B_LV_ATTRIBUTE + 1) * 10)
				ConsoleWrite("+" & $A_ROMList[$B_ROMList][0] & ' -> ' & "(" & $A_LV_ATTRIBUTE[$B_LV_ATTRIBUTE] & ") ===> " & $A_ROMList[$B_ROMList][2] & '+' & ($B_LV_ATTRIBUTE * 10) & @CRLF) ; Debug
			EndIf
		Next
		$A_ROMList[$B_ROMList][3] = $A_ROMList[$B_ROMList][1] & $A_ROMList[$B_ROMList][2]
	Next
;~ 	_ArrayDisplay($A_ROMList, '$A_ROMList Completed') ; Debug
	_ArraySort($A_ROMList, 0, 0, 0, 3)
;~ 	_ArrayDisplay($A_ROMList, '$A_ROMList Sorted') ; Debug
	ProgressOff()
	Return $A_ROMList
EndFunc   ;==>_MOVE_ROM

Func _SUPPR_ROM($V_ROMPath, $I_LV_SUPPRESS, $A_ROMList) ;Définition des ROMs à ne pas garder (Chemin des ROMs, Indexe de la LV des attributs non conservé, Array des ROMs)
	ProgressOn(_MultiLang_GetText("prbr_suppr_rom_title"), "", "0%")
	$A_LV_SUPPRESS = _GUIListViewEx_ReturnArray($I_LV_SUPPRESS)
	For $B_ROMList = 0 To UBound($A_ROMList) - 1
		$V_ProgressPRC = Round(($B_ROMList * 100) / (UBound($A_ROMList) - 1))
		ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_suppr_rom_progress") & $V_ProgressPRC & "%")
		For $B_LV_SUPPRESS = 0 To UBound($A_LV_SUPPRESS) - 1
			If StringInStr($A_ROMList[$B_ROMList][0], "(" & $A_LV_SUPPRESS[$B_LV_SUPPRESS] & ")") > 0 Then
				ConsoleWrite($V_ROMPath & $A_ROMList[$B_ROMList][0] & ' -> ' & $V_ROMPath & "MOVED\" & $A_ROMList[$B_ROMList][1] & StringRight($A_ROMList[$B_ROMList][0], 4) & @CRLF) ; Debug
				$A_ROMList[$B_ROMList][2] = "SUPPR"
			EndIf
		Next
	Next
;~ 	_ArrayDisplay($A_ROMList, '$A_ROMList Completed') ; Debug
	ProgressOff()
	Return $A_ROMList
EndFunc   ;==>_SUPPR_ROM

Func _CLEAN_ROM($V_ROMPath, $A_ROMList, $TMP_Action = 0) ;Nettoyage des ROMs (Chemin des ROMs, Array des ROMs, Action = 0-Simulation;1-Nettoyage)
	ProgressOn(_MultiLang_GetText("prbr_clean_rom_title"), "", "0%")
	$A_ROMList_CLEAN = $A_ROMList
	For $B_ROMList = UBound($A_ROMList_CLEAN) - 1 To 0 Step -1
		$V_ProgressPRC = Round((((UBound($A_ROMList_CLEAN) - 1) - $B_ROMList) * 100) / (UBound($A_ROMList_CLEAN) - 1))
		ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_clean_rom_progress0") & $V_ProgressPRC & "%")
		If $A_ROMList_CLEAN[$B_ROMList][2] = "SUPPR" Then
			_ArrayDelete($A_ROMList_CLEAN, $B_ROMList)
		EndIf
	Next

	Local $D_ROMList = ObjCreate("Scripting.Dictionary")
	If @error Then Exit MsgBox(17, '', 'ERROR getting object')

	For $B_ROMList = 0 To UBound($A_ROMList_CLEAN) - 1
		$V_ProgressPRC = Round(($B_ROMList * 100) / (UBound($A_ROMList_CLEAN) - 1))
		ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_clean_rom_progress1") & $V_ProgressPRC & "%")
		If Not $D_ROMList.exists($A_ROMList_CLEAN[$B_ROMList][1]) Then
			$D_ROMList.add($A_ROMList_CLEAN[$B_ROMList][1], $A_ROMList_CLEAN[$B_ROMList][0])
		Else
			$D_ROMList.item($A_ROMList_CLEAN[$B_ROMList][1]) = $A_ROMList_CLEAN[$B_ROMList][3]
		EndIf
	Next

	Local $K_ROMList = $D_ROMList.keys(), $A_ROMList_TEMP[$D_ROMList.count()][2]
	For $B_ROMList = 0 To UBound($K_ROMList) - 1
		$V_ProgressPRC = Round(($B_ROMList * 100) / (UBound($K_ROMList) - 1))
		ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_clean_rom_progress3") & $V_ProgressPRC & "%")
		$A_ROMList_TEMP[$B_ROMList][0] = $D_ROMList.Item($K_ROMList[$B_ROMList])
		$A_ROMList_TEMP[$B_ROMList][1] = $K_ROMList[$B_ROMList]
	Next

;~ 	_ArrayDisplay($A_ROMList, "$A_ROMList");Debug
;~ 	_ArrayDisplay($A_ROMList_TEMP, "$A_ROMList_TEMP");Debug
	Dim $A_ROMList_SIMUL[UBound($A_ROMList)][2]
	For $B_ROMList = 0 To UBound($A_ROMList) - 1
		$V_ProgressPRC = Round(($B_ROMList * 100) / (UBound($A_ROMList) - 1))
		ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_clean_rom_progress3") & $V_ProgressPRC & "%")
		For $B_ROMList_TEMP = 0 To UBound($A_ROMList_TEMP) - 1
			If $A_ROMList[$B_ROMList][0] = $A_ROMList_TEMP[$B_ROMList_TEMP][0] Then $A_ROMList[$B_ROMList][2] = "KEEP"
		Next
	Next

	For $B_ROMList_SIMUL = 0 To UBound($A_ROMList_SIMUL) - 1
		$V_ProgressPRC = Round(($B_ROMList_SIMUL * 100) / (UBound($A_ROMList_SIMUL) - 1))
		ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_clean_rom_progress4") & $V_ProgressPRC & "%")
		Select
			Case $A_ROMList[$B_ROMList_SIMUL][2] = "KEEP"
				$A_ROMList_SIMUL[$B_ROMList_SIMUL][0] = "OK"
				$A_ROMList_SIMUL[$B_ROMList_SIMUL][1] = $A_ROMList[$B_ROMList_SIMUL][0]
			Case Else
				$A_ROMList_SIMUL[$B_ROMList_SIMUL][0] = "KO"
				$A_ROMList_SIMUL[$B_ROMList_SIMUL][1] = $A_ROMList[$B_ROMList_SIMUL][0]
		EndSelect
	Next
;~ 	_ArrayDisplay($A_ROMList_SIMUL, "$A_ROMList_SIMUL");Debug
	_FileWriteFromArray($path_SIMUL, $A_ROMList_SIMUL, 1)
;~ 	_ArrayDisplay($A_ROMList_TEMP, "result");Debug
	If $TMP_Action = 1 Then
		For $B_ROMList = 0 To UBound($A_ROMList_TEMP) - 1
			$V_ProgressPRC = Round(($B_ROMList * 100) / (UBound($A_ROMList_TEMP) - 1))
			ProgressSet($V_ProgressPRC, _MultiLang_GetText("prbr_clean_rom_progress2") & $V_ProgressPRC & "%")
			FileMove($V_ROMPath & $A_ROMList_TEMP[$B_ROMList][0], $V_ROMPath & "CLEAN_ROM\", BitOR($FC_OVERWRITE, $FC_CREATEPATH))
		Next
		ProgressOff()
	Else
		ProgressOff()
		; Display the file.
		ShellExecute($path_SIMUL)
	EndIf
EndFunc   ;==>_CLEAN_ROM

Func _LANG_LOAD($LANG_DIR, $user_lang) ;Chargement de la langue (Chemin des fichiers de langues, Id de la langue)
	;Create an array of available language files
	; ** n=0 is the default language file
	; [n][0] = Display Name in Local Language (Used for Select Function)
	; [n][1] = Language File (Full path.  In this case we used a $LANG_DIR
	; [n][2] = [Space delimited] Character codes as used by @OS_LANG (used to select correct lang file)
	Local $LANGFILES[2][3]

	$LANGFILES[0][0] = "English (US)" ;
	$LANGFILES[0][1] = $LANG_DIR & "\URC-ENGLISH.XML"
	$LANGFILES[0][2] = "0409 " & _ ;English_United_States
			"0809 " & _ ;English_United_Kingdom
			"0c09 " & _ ;English_Australia
			"1009 " & _ ;English_Canadian
			"1409 " & _ ;English_New_Zealand
			"1809 " & _ ;English_Irish
			"1c09 " & _ ;English_South_Africa
			"2009 " & _ ;English_Jamaica
			"2409 " & _ ;English_Caribbean
			"2809 " & _ ;English_Belize
			"2c09 " & _ ;English_Trinidad
			"3009 " & _ ;English_Zimbabwe
			"3409" ;English_Philippines

	$LANGFILES[1][0] = "Français" ; French
	$LANGFILES[1][1] = $LANG_DIR & "\URC-FRENCH.XML"
	$LANGFILES[1][2] = "040c " & _ ;French_Standard
			"080c " & _ ;French_Belgian
			"0c0c " & _ ;French_Canadian
			"100c " & _ ;French_Swiss
			"140c " & _ ;French_Luxembourg
			"180c" ;French_Monaco

	;Set the available language files, names, and codes.
	_MultiLang_SetFileInfo($LANGFILES)
	If @error Then
		MsgBox(48, "Error", "Could not set file info.  Error Code " & @error)
		Exit
	EndIf
	;Check if the loaded settings file exists.  If not ask user to select language.
	If $user_lang = -1 Then
		;Create Selection GUI
		$user_lang = _LANGUE_SelectGUI($LANGFILES)
		If @error Then
			MsgBox(48, "Error", "Could not create selection GUI.  Error Code " & @error)
			Exit
		EndIf
		IniWrite($PathConfigINI, "LAST_USE", "$user_lang", $user_lang)
	EndIf
	Local $ret = _MultiLang_LoadLangFile($user_lang)
	If @error Then
		MsgBox(48, "Error", "Could not load lang file.  Error Code " & @error)
		Exit
	EndIf
	;If you supplied an invalid $user_lang, we will load the default language file
	If $ret = 2 Then
		MsgBox(64, "Information", "Just letting you know that we loaded the default language file")
	EndIf

	Return $LANGFILES
EndFunc   ;==>_LANG_LOAD

Func _LANGUE_SelectGUI($_gh_aLangFileArray, $default = @OSLang) ;Interface de séléction de la langue (Array des langues, langue par défaut)
	GUISetState(@SW_DISABLE, $F_UniversalCleaner)
	If $_gh_aLangFileArray = -1 Then Return SetError(1, 0, 0)
	If IsArray($_gh_aLangFileArray) = 0 Then Return SetError(1, 0, 0)
	Local $_multilang_gui_GUI = GUICreate(_MultiLang_GetText("win_sel_langue_Title"), 230, 100)
	Local $_multilang_gui_Combo = GUICtrlCreateCombo("(" & _MultiLang_GetText("win_sel_langue_Title") & ")", 8, 48, 209, 25, BitOR($CBS_DROPDOWNLIST, $CBS_AUTOHSCROLL))
	Local $_multilang_gui_Button = GUICtrlCreateButton(_MultiLang_GetText("win_sel_langue_button"), 144, 72, 75, 25)
	Local $_multilang_gui_Label = GUICtrlCreateLabel(_MultiLang_GetText("win_sel_langue_text"), 8, 8, 212, 33)
	;Create List of available languages
	For $i = 0 To UBound($_gh_aLangFileArray) - 1
		GUICtrlSetData($_multilang_gui_Combo, $_gh_aLangFileArray[$i][0], "(" & _MultiLang_GetText("win_sel_langue_Title") & ")")
	Next
	GUISetState(@SW_SHOW)
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case -3, $_multilang_gui_Button
				ExitLoop
		EndSwitch
	WEnd
	Local $_selected = GUICtrlRead($_multilang_gui_Combo)
	GUIDelete($_multilang_gui_GUI)
	For $i = 0 To UBound($_gh_aLangFileArray) - 1
		If StringInStr($_gh_aLangFileArray[$i][0], $_selected) Then
			GUISetState(@SW_ENABLE, $F_UniversalCleaner)
			WinActivate($F_UniversalCleaner)
			Return StringLeft($_gh_aLangFileArray[$i][2], 4)
		EndIf
	Next
	GUISetState(@SW_ENABLE, $F_UniversalCleaner)
	WinActivate($F_UniversalCleaner)
	Return $default
EndFunc   ;==>_LANGUE_SelectGUI

