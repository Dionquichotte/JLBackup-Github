; *******************************************************************************************************************************************************
; JL backup, program for moving joblistfiles on Tecan Twin (Green and Blue) system
; Sanquin IPB Laboratory of Biologicals
;
; version 1.3, 20 august 2014 by Dion Methorst
; version 1.4, 03 february 2015
; version 1.5, 17 september 2015
; version 1.6, 23 september 2015
; version 1.7, 23 september 2015
; version 2.0, 12 januari 2016
;
; *******************************************************************************************************************************************************
;
; changelog:
; v1.1		added 5 more EL folders for new CLIS worklist , these numbers are arbitrary and will be replaced for real world worklists later
; v1.2		line 104 and 119, Enbrel changed worklistnumber 699 to 923
; v1.3		line 109 & 123, arrays with worklistnumbers for EVO $Blue[7] and $Green[7] changed to Golimumab
; 			line 85-100, size of arrays $Blue and $Green now set as amount of ELn folders counted in C:\Apps\EVO\Job\
; 			line 85-100, added errorhandling to check for existing ELn folders
;			added checking of existing files in joblist folder, existing joblist files in folder are not overwritten.
;			adjusted messagebox icon and messages for clarity
; v1.4		reading EL folders and joblistfiles from ini file:
; v1.5		Added RenameJoblist function: Resets the count of the joblists in each ELx folder
; v1.6		Rename Joblist function opening joblist folder to check joblists.
;			Transparent GUI
; v1.7		added buttons to open EVO file folders
; v2.0		Tab menu's for functions and shortcuts
;			Shortcuts assigned in ini file
;			Removed obsolete 'Move' and 'Reset' joblist function
;			Gui always on top: Guicreate... $WS_EX_TOPMOST
;
; *******************************************************************************************************************************************************
;
; Upon execution of the  main GUI window with 3 buttons pops up in the first tab:
;
; Each of the buttons executes a function of the same name as the button as described below.
;
; set joblistcount from 1 to 5
; delete joblists
; exit
;
; *******************************************************************************************************************************************************
;
; function RenameJoblist()
;
; Reads all joblists in each C:apps/EVO/JOB/ELx folder and renames them as ASSAYnmbr_EVOnumber_Y.twl where Y = 1 to 5.
; Normally only 5 joblists are present in all ELx folders combined,
; Nevertheless all joblists are renamed in each ELx folder, but when the number of joblists exceeds five, a messagebox will display the number of joblists present
; The C;apps/EVO/JOB folder will be opened and another message instructs the operator to check if the renaming of the joblists was succesful
; (This script cannot distinguish between the timestamp of CLIS generated joblists, nor of the oprators intentions and choices &
; as such these will have to be addressed by the operator or changes to the CLIS database will have to be made)
;
; funcrtion DeleteJoblist()
;
; script to delete ALL joblists from C:\APPS\EVO\job\ELn folders
; if not backupped beforehand, all files will be lost forevermore :)
; therefore the main() gui pops a message if you're sure about throwing away your joblists
;
;			FileDelete("C:\apps\EVO\job\EL" & $EL) ; deletes *.* in folder
;
; the array is put in a loop which designates a number to the ELx [1-x] folders
; the loop is executed x-1 times, thus EL1, EL2... EL5 are named
; This loop can also be extended up to as many ELx folders as you wish
;
; *******************************************************************************************************************************************************
;
; The second tab of the main GUI has 5 buttons on the left side and 5 buttons on the right side
;
; These buttons can be designated in the JLbackup.ini file
; Jlbackup.ini   >> a default JLBackup is created after deletion of the ini file.
; JL Backup.ini file location is the same as the script location
;
; [ShortcutsL] section for ButtonL1 to ButtonL5 and [ShortcutsR] section for ButtonR1 to ButtonR5,
; followed by five keys for the button names with accompanying values for the paths to open
;
; *******************************************************************************************************************************************************
;
; the default values are shown in the ini file example:
;
;[Blue]
;EL1= \682_2_		 infliximab = remicade
;EL2= \923_2_		enbrel = etanercept
;EL3 = \683_2_		adalimumab = humera
;EL4 = \697_2_		rituximab = mapthera
;EL5 = \800_2_		trastuzumab = herceptin
;EL6 = \856_2_		tocilizumab
;EL7 = \922_2_		golimumab
;EL8 = \999_2_		bloedspot ADA
;EL9 = \998_2_		bloedspot ETN
;EL10 = \997_2_
;
;[Green]
;EL1 = \682_1_
;EL2 = \923_1_
;EL3 = \683_1_
;EL4 = \697_1_
;EL5 = \800_1_
;EL6 = \856_1_
;EL7 = \922_1_
;EL8 = \999_1_
;EL9 = \998_1_
;
;[ShortcutsL]
;JOB=C:\APPS\EVO\JOB
;TPL=C:\APPS\EVO\TPL
;ASC=C:\APPS\EVO\ASC
;TPLASC=C:\APPS\EVO\TPLASC
;Archief=C:\APPS\EVO\Archief
;
;[ShortcutsR]
;EVOware=C:\ProgramFiles (x86)
;AppData=C:\Users\Administrator\AppData
;ProgramData=C:\ProgramData\Tecan\Evoware
;Documents=C:\Users\Administrator\Documents
;ButtonR5=C:\ProgramData\Tecan\Evoware
;
;******************************************************************************************************************************************************
; Start of script

#include <Array.au3>
#include <file.au3>
#include <Date.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <WinAPI.au3>
#include <WindowsConstants.au3>
#include <Constants.au3>
#include <Math.au3>
#include <ButtonConstants.au3>
#include <TabConstants.au3>

If not FileExists(@ScriptDir & "\Jlbackup.ini") then
	Local $BlueAssay = "EL1=\682_2_" & @CRLF &"EL2=\923_2_" & @CRLF &"EL3=\683_2_" & @CRLF &"EL4=\697_2_" & @CRLF &"EL5=\800_2_" _
	& @CRLF &"EL6=\856_2_" & @CRLF &"EL7=\922_2_" & @CRLF &"EL8=\999_2_" & @CRLF &"EL9=\998_2_"
	Local $GreenAssay = "EL1=\682_1_" & @CRLF &"EL2=\923_1_" & @CRLF &"EL3=\683_1_" & @CRLF &"EL4=\697_1_" & @CRLF &"EL5 =\800_1_" _
	& @CRLF &"EL6=\856_1_" & @CRLF &"EL7=\922_1_" & @CRLF &"EL8=\999_1_" & @CRLF &"EL9=\998_1_"
	Local $ShortLC =  "JOB=C:\APPS\EVO\JOB"& @CRLF & "TPL=C:\APPS\EVO\TPL" & @CRLF & "ASC=C:\APPS\EVO\ASC" _
	& @CRLF & "TPLASC=C:\APPS\EVO\TPLASC" & @CRLF & "Archief=C:\APPS\EVO\Archief"
	Local $ShortRC =  "EVOware=C:\ProgramFiles (x86)"& @CRLF & "AppData=C:\Users\Administrator\AppData" _
	& @CRLF & "ProgramData=C:\ProgramData\Tecan\Evoware" & @CRLF & "Documents=C:\Users\Administrator\Documents" & @CRLF & "PC=C:\"

	IniWriteSection(@ScriptDir & "\Jlbackup.ini", "Blue", $BlueAssay)
	IniWriteSection(@ScriptDir & "\Jlbackup.ini", "Green", $GreenAssay)
	IniWriteSection(@ScriptDir & "\Jlbackup.ini", "ShortcutsL", $ShortLC)
	IniWriteSection(@ScriptDir & "\Jlbackup.ini", "ShortcutsR", $ShortRC)
	Endif

$Blue = IniReadSection(@ScriptDir & "\Jlbackup.ini", "Blue")
$Green = IniReadSection(@ScriptDir & "\Jlbackup.ini", "Green")
$ShortcutsL = IniReadSection(@ScriptDir & "\Jlbackup.ini", "ShortcutsL")
$ShortcutsR = IniReadSection(@ScriptDir & "\Jlbackup.ini", "ShortcutsR")

;_ArrayDisplay($Blue)
;_ArrayDisplay($Green)
;_ArrayDisplay($ShortcutsL)
;_ArrayDisplay($ShortcutsR)

;================== Start Main() ==============================================================================================================
; GUI Creation
GUISetFont(9, 500, 2, 45)
Global const $JLbu = GuiCreate("EVO Joblist Backup", 255, 205, -1, -1, -1, BitOr($WS_EX_TOOLWINDOW, $WS_EX_LAYERED, $WS_EX_TOPMOST)) ;BitOr($WS_EX_TRANSPARENT, $WS_EX_TOOLWINDOW , $WS_EX_LAYERED))
GuiSetIcon("B:\Programmeren\programmeren_Dion\TECAN\EVO joblist backup Allergie\icons\JLbackup2.ico", 0)
;DllCall("user32.dll", "int", "AnimateWindow", "hwnd", $JLbu, "int", 1000, "long", 0x00080000) ; fade-in

$Tab1 = GUICtrlCreateTab(8, 8, 240, 190)
$TabSheet1 = GUICtrlCreateTabItem("JoblistBackup")
$RenameJoblist = GUICtrlCreateButton("Set Joblists 1 to 5",15,35,226,45)
$DeleteButton = GUICtrlCreateButton("Delete Joblists",15,90,225,45)
$ExitButton = GUICtrlCreateButton("Exit",15,145,225,45)
	;OLD
	;$MoveJoblist = GUICtrlCreateButton("Move Joblists",15,65,165,25)
	;$ResetJoblist = GUICtrlCreateButton("Reset all Joblists",15,95,165,25)
	;$DeleteButton = GUICtrlCreateButton("Delete Joblists",15,125,165,25)
	;$ExitButton = GUICtrlCreateButton("Exit",15,155,165,25)

$TabSheet2 = GUICtrlCreateTabItem("Shortcuts")
$ButtonL1 = GUICtrlCreateButton($ShortcutsL[1][0],15,37,75,25)
$ButtonL2 = GUICtrlCreateButton($ShortcutsL[2][0],15,69,75,25)
$ButtonL3 = GUICtrlCreateButton($ShortcutsL[3][0],15,101,75,25)
$ButtonL4 = GUICtrlCreateButton($ShortcutsL[4][0],15,133,75,25)
$ButtonL5 = GUICtrlCreateButton($ShortcutsL[5][0],15,165,75,25)

$ButtonR1 = GUICtrlCreateButton($ShortcutsR[1][0], 163, 37, 75, 25)
$ButtonR2 = GUICtrlCreateButton($ShortcutsR[2][0], 163, 69, 75, 25)
$ButtonR3 = GUICtrlCreateButton($ShortcutsR[3][0], 163, 101, 75, 25)
$ButtonR4 = GUICtrlCreateButton($ShortcutsR[4][0], 163, 133, 75, 25)
$ButtonR5 = GUICtrlCreateButton($ShortcutsR[5][0], 163, 165, 75, 25)
GUICtrlCreateTabItem("")

; Close Group
GUICtrlCreateGroup("",-99,-99,1,1)
_WinAPI_SetLayeredWindowAttributes($Tab1, 0xABCDEF);, 125)
_WinAPI_SetLayeredWindowAttributes($JLbu, 0xABCDEF);, 125)
GUISetBkColor(0xABCDEF)
; Show windows with buttons
GuiSetState(@SW_SHOW)
;GUISetState()

; Continuous Loop to check for GUI Events, upon event the corresponding command or function is exectuted
While 1
$guimsg = GUIGetMsg()
	Select
		 Case $guimsg = $GUI_EVENT_CLOSE
			ExitLoop
		 Case $guimsg = $RenameJoblist
			RenameJoblist()
			   MsgBox(0,"Joblist Count","Joblists Count Done!")
				  Local $iPID = Run("explorer.exe " & "C:\Apps\EVO\Job")
				  WinWait("[CLASS:explorer]", "", 1)
				  Sleep(10)
				  ProcessClose($iPID)
			   MsgBox ($MB_ICONINFORMATION + $MB_TOPMOST + $MB_SETFOREGROUND, "", "Check Joblists!",5)
		 Case $guimsg = $ButtonL1
			Local $iPID = Run("explorer.exe " & $ShortcutsL[1][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)
		 Case $guimsg = $ButtonL2
			Local $iPID = Run("explorer.exe " & $ShortcutsL[2][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)
		 Case $guimsg = $ButtonL3
			Local $iPID = Run("explorer.exe " & $ShortcutsL[3][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)
		 Case $guimsg = $ButtonL4
			Local $iPID = Run("explorer.exe " & $ShortcutsL[4][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)
		 Case $guimsg = $ButtonL5
			Local $iPID = Run("explorer.exe " & $ShortcutsL[5][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)

		;buttons as defined by shortcuts section in ini file:
		Case $guimsg =$ButtonR1
			Local $iPID = Run("explorer.exe " & $ShortcutsR[1][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)
		Case $guimsg = $ButtonR2
			Local $iPID = Run("explorer.exe " & $ShortcutsR[2][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)
		Case $guimsg = $ButtonR3
			Local $iPID = Run("explorer.exe " & $ShortcutsR[3][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)
		Case $guimsg = $ButtonR4
			Local $iPID = Run("explorer.exe " & $ShortcutsR[4][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)
		Case $guimsg = $ButtonR5
			Local $iPID = Run("explorer.exe " & $ShortcutsR[5][1])
			WinWait("[CLASS:explorer]", "", 1)
			Sleep(10)
			ProcessClose($iPID)
		Case $guimsg = $ExitButton
			Exit
		 Case $guimsg = $DeleteButton
			Dim $iMsgBoxAnswer
			$iMsgBoxAnswer = MsgBox(52,"DELETE Joblists"," Weet u dit heel erg zeker?")
			Select
				Case $iMsgBoxAnswer = 6 ;Yes
					DeleteJoblist()
					MsgBox(64,"Joblist DELETE","Joblists Deleted")
				Case $iMsgBoxAnswer = 7 ;No
					MsgBox(64,"Joblist DELETE","Joblists NOT Deleted")
			EndSelect
	EndSelect
Wend
;================== END Main() ==============================================================================================================
;================== START FUNCTION Delete Joblist ===========================================================================================
Func DeleteJoblist()

$max = _Max ($Green[0][0], $Blue[0][0])

For $EL = 1 to $max
			FileDelete("C:\apps\EVO\job\EL" & $EL) ; deletes *.* in folder
Next

EndFunc
;================== END FUNCTION Delete Joblist ================================================================================================
;================== START FUNCTION Rename Joblist ==============================================================================================
Func RenameJoblist()

 $aFileList = _FileListToArray("C:\apps\EVO\job\", "*")
 $aFileList2 = _FileListToArray("C:\apps\EVO\job\", Default, Default, True)
_arraydisplay($aFilelist2)

$JLcount = 1
For $i = 1 to Ubound($aFileList2)-1
	   $aJobList = _FileListToArray("C:\apps\EVO\job\"& $aFileList[$i] & "\" , "*.twl")

	;_arraysort($aJoblist)
	;_arraydisplay($aJobList)
	  for $j = 1 to Ubound($aJobList)-1
		 $pos =stringinstr($aJobList[$j], ".", 0,1)-1
		 filemove($aFileList2[$i] & "\" & $aJobList[$j] , $aFileList2[$i] & "\" & StringReplace($aJobList[$j], $pos, $j & ".twl"), 1 +8)
		 $JLcount = $JLcount + 1
		next
	  ;msgbox(0, "", $JLcount)
Next

if $JLcount >5 then msgbox($MB_ICONWARNING, "Opgepast!", "Er zijn meer dan " & $JLcount & " joblists aanwezig" & @CRLF & "Dat zijn er meer dan 5!")

EndFunc
;================== END FUNCTION Rename Joblist ================================================================================================