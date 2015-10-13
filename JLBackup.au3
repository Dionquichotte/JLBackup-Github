; *******************************************************************************************************************************************************
; JL backup, program for moving joblistfiles on Tecan Twin (Green and Blue) system
; Sanquin IPB Laboratory of Biologicals
;
; version 1.3, 20 august 2014 by Dion Methorst
;
; changelog:
; v1.1	 added 5 more EL folders for new CLIS worklist , these numbers are arbitrary and will be replaced for real world worklists later
; v1.2	 line 104 and 119, Enbrel changed worklistnumber 699 to 923
; v1.3	 line 109 & 123, arrays with worklistnumbers for EVO $Blue[7] and $Green[7] changed to Golimumab
; 		 line 85-100, size of arrays $Blue and $Green now set as amount of ELn folders counted in C:\Apps\EVO\Job\
; 		 line 85-100, added errorhandling to check for existing ELn folders
;		 added checking of existing files in joblist folder, existing joblist files in folder are not overwritten.
;		 adjusted messagebox icon and messages for clarity
;
;
; Future changes: (?)open for debate... sort of... :)
; moving joblist array to ini file for manual setting of worklists, lims-number and position in array == setting of ELn folder
;
;
; *******************************************************************************************************************************************************
;
; Upon execution of the  main GUI window with 4 buttons pops up:
;
; move joblists
; reset joblists
; exit
; delete joblists
;
; each of the buttons executes a function of the same name as the button as described below.
;
; function MoveJoblist()
;
; script to move Tecan Freedom joblist files
; The script executes the following procedure:
;
;			FileMove("C:\APPS\EVO\job\EL" & $ELB & $Blue[$i] & $File & ".twl", "C:\;APPS\EVO\JLbackup" & $Blue[$i] & $File & ".twl", 1 + 8)
;			FileCopy("C:\APPS\EVO\JLbackup" & $Blue[$i] & $File & ".twl", "C:\APPS\E;VO\archief" & $Blue[$i] & $File & "_" & $filedatum & ".twl", 1 + 8)
;			FileSetTime("C:APPS\EVO\JLbackup" & $Blue[$i] & $File  & ".twl", "", 1)
;
; $f is amount of ELn folders in C:\APPS\EVO\JOB\
; Files are moved from C:\APPS\EVO\JOB\ELn (n= 1 tot $f)
; Files are moved to C:\APPS\EVO\JLbackup
; The original files are copied from C:\APPS\EVO\JLbackup\ to C:\APPS\EVO\Archief,
; time of copying is set to JLBackup folder and all files as file attribute "date created"
; date and timestamp are inserted into the filename of files in \ARCHIEF
;
; function ResetJoblist()
;
; script to reset Tecan Freedom joblist files from JLbackup folder
; this function reverses the MoveJoblist() and puts the files back into the EL Joblist folders
; existing files in the joblistfolders are checked out first in order that existing joblist files are NOT overwritten!!!
;
;			; move files $Blue, 1st check if file already in joblist folder, if NOT then joblists ar moved to backup folder and archief, timestamp added.
;			; existing files are not overwritten
;			If FileExists("C:\APPS\EVO\Job\EL" & $ELB  & $Blue[$i] & $File & ".twl") Then
;				MsgBox(4096, "JLBackup", "C:\APPS\EVO\Job"  & $Blue[$i] & $File & ".twl" & "already exists in joblist folder!" &  @CRLF & _
;				"This file is not resetted and will be kept in C:\apps\EVO\JLBackup")
;			Else
;				FileMove("C:\APPS\EVO\JLbackup"  & $Blue[$i] & $File & ".twl", "C:\APPS\EVO\job\EL" & $ELB & $Blue[$i] & $File & ".twl", 0 + 8)
;				FileSetTime("C:\APPS\EVO\job\EL" & $ELB & $Blue[$i] & $File  & ".twl", "", 1)
;			EndIf
;
; The script executes the following procedure:
; Files are copied from C:\APPS\EVO\JLbackup
; Files are copied to C:\APPS\EVO\JOB\ELn (n= 1 tot $f)
;
; funcrtion DeleteJoblist()
;
; script to delete ALL joblists from C:\APPS\EVO\job\ELn folders
; if not backupped beforehand, all files will be lost forevermore :)
; therefore the main() gui pops a messagebox whether you're sure about throwing away your joblists
;
;			FileDelete("C:\apps\EVO\job\EL" & $EL)
;
; joblist IDs are defined in an array of 10 variables
; [0-4][10-14] for EVO Blue and [5-9][15-20] for EVO Green
;
; This array can be expanded to as many files as you wish : )
; the array is put in a loop which designates a number to the ELx [1-5] folders
; the loop is executed 5 times, thus EL1, EL2... EL5 are named
; This loop can also be extended up to as many ELx folders as you wish
;
;******************************************************************************************************************************************************

; Start of script

; library inclusion
#include <Array.au3>
#include <file.au3>
#include <Date.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <WinAPI.au3>
#include <Constants.au3>

; Joblist backup data array
Global $Blue[10] ;$Blue[$i]

;count folders in C:\apps\evo\job\ + errorhandling if none at all
Local $FileList = _FileListToArray("C:\APPS\EVO\Job")
	If @error = 1 Then
		MsgBox(0, "", "No Folders Found.")
		Exit
	EndIf
	If @error = 4 Then
		MsgBox(0, "", "No Files Found.")
		Exit
	EndIf
;_ArrayDisplay($FileList, "$FileList") ;for testing purposes

; Retrieve amount of ELn folders for setting size of joblist array
$f = $Filelist[0]

; joblist array EVO Blue
Global $Blue[$f+1] ;$Blue[$i]

	$Blue[0] = "EVO Blue" ;IP adress/Sanquin number? or username? get it from logfile or ini
	$Blue[1] = "\682_2_" ; infliximab = remicade
	$Blue[2] = "\923_2_" ; enbrel = etanercept
	$Blue[3] = "\683_2_" ; adalimumab = humera
	$Blue[4] = "\697_2_" ; rituximab = mapthera
	$Blue[5] = "\800_2_" ; trastuzumab = herceptin
	$Blue[6] = "\856_2_" ; tocilizumab
	$Blue[7] = "\922_2_" ; golimumab
	;$Blue[8] = "\827_2_" ; natalizumab
	;$Blue[9] = "\846_2_" ; abatacept
	;$Blue[10] = "\xxx_2_"

; joblist array EVO Green
Global $Green[$f+1] ;$Green[$p]
	$Green[0] = "EVO Green"
	$Green[1] = "\682_1_"
	$Green[2] = "\923_1_"
	$Green[3] = "\683_1_"
	$Green[4] = "\697_1_"
	$Green[5] = "\800_1_"
	$Green[6] = "\856_1_"
	$Green[7] = "\922_1_"
	;$Green[8] = "\827_1_"
	;$Green[9] = "\846_1_" ; abatacept
	;$Green[10] = "\xxx_1_"

;_ArrayDisplay($Blue)
;_ArrayDisplay($Green)
;================== START FUNCTION Move Joblist=======================================================================================================

; function for use in GUI, this script transfers all  Joblist files to the C:\apps\EVO\archief folder
Func MoveJoblist()
dim $EL
dim $ELB
dim $ELG
dim $File = ""
dim $filedatum

$filedatum = @mday & @mon & @year & @hour & @min & @sec

For $EL = 1 to Ubound($Blue)-1 ; loop to move EVO blue joblists corresponding to array

			$ELB = $EL	; may seem obsolete but is for future use when EVO BLue and Green 'grow out of sync'
			$ELG = $EL	; and the loop has to be divided in two seperate parts
			$i = $ELB
			$p = $ELG

			For $File = 1 to 5

			; move files $Blue, joblists ar moved to backup folder and archief, timestamp added
			FileMove("C:\APPS\EVO\job\EL" & $ELB & $Blue[$i] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $Blue[$i] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $Blue[$i] & $File & ".twl", "C:\APPS\EVO\archief" & $Blue[$i] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $Blue[$i] & $File  & ".twl", "", 1)

			; move files $Green, joblists ar moved to backup folder and archief, timestamp added
			FileMove("C:\APPS\EVO\job\EL" & $ELG & $Green[$p] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $Green[$p] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $Green[$p] & $File & ".twl", "C:\APPS\EVO\archief" & $Green[$p] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $Green[$p] & $File  & ".twl", "", 1)

			next

Next

EndFunc
;================== END FUNCTION move Joblist =====================================================================================================

;================== START FUNCTION Reset Joblist ========================================================================================================
Func ResetJoblist()

dim $EL
dim $ELB
dim $ELG
dim $File = ""
dim $filedatum

For $EL = 1 to Ubound($Blue)-1 ; loop to move EVO blue joblists corresponding to array

			$ELB = $EL	; may seem obsolete but is for future use when EVO BLue and Green 'grow out of sync'
			$ELG = $EL	; and the loop has to be divided in two seperate parts
			$i = $ELB
			$p = $ELG

		For $File = 1 to 5

			; move files $Blue, 1st check if file already in joblist folder, if NOT then joblists ar moved to backup folder and archief, timestamp added.
			; existing files are not overwritten
			If FileExists("C:\APPS\EVO\Job\EL" & $ELB  & $Blue[$i] & $File & ".twl") Then
				MsgBox(4096, "JLBackup", "C:\APPS\EVO\Job"  & $Blue[$i] & $File & ".twl" & "already exists in joblist folder!" &  @CRLF & _
				"This file is not resetted and will be kept in C:\apps\EVO\JLBackup")
			Else
				FileMove("C:\APPS\EVO\JLbackup"  & $Blue[$i] & $File & ".twl", "C:\APPS\EVO\job\EL" & $ELB & $Blue[$i] & $File & ".twl", 0 + 8)
				FileSetTime("C:\APPS\EVO\job\EL" & $ELB & $Blue[$i] & $File  & ".twl", "", 1)
			EndIf

			; move files $Green, as above
			If FileExists("C:\APPS\EVO\Job\EL" & $ELG & $Green[$p] & $File & ".twl") Then
				MsgBox(4096, "JLBackup", "C:\APPS\EVO\Job"  & $Green[$i] & $File & ".twl" & "already exists in joblist folder!" &  @CRLF & _
				"This file is not resetted and will be kept in C:\apps\EVO\JLBackup")
			ELse
				FileMove("C:\APPS\EVO\JLbackup" & $Green[$p] & $File & ".twl", "C:\APPS\EVO\job\EL" & $ELG & $Green[$p] & $File & ".twl", 0 + 8)
				FileSetTime("C:\APPS\EVO\job\EL" & $ELG & $Green[$p] & $File  & ".twl", "", 1)
			EndIf

		next

Next

EndFunc
;================== END FUNCTION Reset Joblist ================================================================================================

;================== START FUNCTION Delete Joblist ========================================================================================================
Func DeleteJoblist()

dim $EL
dim $File = ""
; n = 1/5, ELn folders: EL1, EL2, etc
; there is a maximum of 5 joblists per EL folder

For $EL = 1 to Ubound($BLue)-1
			;For $File = 1 to 5
			FileDelete("C:\apps\EVO\job\EL" & $EL)
			;Next
Next

EndFunc
;================== END FUNCTION Delete Joblist ================================================================================================

;================== Start Main() ==============================================================================================================

; GUI Creation
GuiCreate("JL Backup", 200, 200)
GUISetBkColor(0xE0FFFF)
GUISetFont(9, 500, 2, 45)
GuiCreate("EVO Joblist Backup", 200, 150)
GuiSetIcon("C:\Laboratorium\programmeren\TECAN\EVO joblist backup Allergie\icon\JLbackup2.ico", 0)

; Button advanced, menu
$MoveJoblist = GUICtrlCreateButton("Move Joblists",20,15,165,25)
$ResetJoblist = GUICtrlCreateButton("Reset all Joblists",20,45,165,25)
$ExitButton = GUICtrlCreateButton("Exit",20,75,165,25)
$DeleteButton = GUICtrlCreateButton("Delete Joblists",20,105,165,25)

; Close Group
GUICtrlCreateGroup("",-99,-99,1,1)

; Show windows with buttons
GuiSetState(@SW_SHOW)
;GUISetState()

; Continuous Loop to check for GUI Events, upon event the corresponding command or function is exectuted
While 1
$guimsg = GUIGetMsg()
	Select
		Case $guimsg = $GUI_EVENT_CLOSE
			ExitLoop
		Case $guimsg = $MoveJoblist
			MoveJoblist()
			MsgBox(0,"Move Joblist","Joblists moved to C:\APPS\EVO\JLbackup")
		;Exit
		Case $guimsg = $ResetJoblist
			Resetjoblist()
			MsgBox(0,"Reset Joblist","Joblists resetted to C:\APPS\EVO\Job")
		;Exit
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
					;Exit
					MsgBox(64,"Joblist DELETE","Joblists NOT Deleted")
			EndSelect
	EndSelect
Wend

;================== END Main() ==============================================================================================================