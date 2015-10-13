; JL backup, program for moving joblistfiels on Tecan Twin (Green and Blue) system
; version 1.0, 16 januari 2012 by Dion Methorst


; function MoveJoblist()

; script to move Tecan Freedom joblist files
; The script executes the following procedure:
; Files are moved from C:\APPS\EVO\JOB\ELn (n= 1 tot 5)
; Files are moved to C:\APPS\EVO\JLbackup
; The original files are copied from C:\APPS\EVO\JLbackup\ to C:\APPS\EVO\Archief
; and underway date and timestamp are inserted into the filename


; function ResetJoblist()

; script to reset Tecan Freedom joblist files from JLbackup folder
; this function reverses the MoveJoblist() and puts the files back into the EL Joblist folders
; existing files in the joblistfolders are not overwritten
; The script executes the following procedure:
; Files are copied from C:\APPS\EVO\JLbackup
; Files are copied to C:\APPS\EVO\JOB\ELn (n= 1 tot 5)
; Date & timestamp are added to secure a unique file-ID
 
; funcrtion DeleteJoblist()

; script to delete ALL joblists from C:\APPS\EVO\job\ELn folders
; if not backupped beforehand, all files will be lost


; joblist IDs are defined in an array of 10 variables
; [0-4] for EVO Blue and [5-9] for EVO Green

; This array can be expanded to as many files as you wish : )
; the array is put in a loop which designates a number to the ELx [1-5] folders
; the loop is executed 5 times, thus EL1, EL2... EL5 are named
; This loop can also be extended up to as many ELx folders as you wish


; library inclusion
#include <Array.au3>
#include <file.au3>
#include <Date.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <WinAPI.au3>
#include <Constants.au3>


;================== START FUNCTION Move Joblist=======================================================================================================

; function for use in GUI, this script transfers all  Joblist files to the C:\apps\EVO\archief folder
Func MoveJoblist()

dim $EL
dim $File = ""
dim $filedatum

$filedatum = @mday & @mon & @year & @hour & @min & @sec

; msgbox(0 , "filedatum", $filedatum)

Global $JL[10]

; joblist files EVO Blue
$JL[0] = "\682_1_"
$JL[1] = "\699_1_"
$JL[2] = "\683_1_"
$JL[3] = "\697_1_"
$JL[4] = "\800_1_"

; joblist files EVO Green
$JL[5] = "\682_2_"
$JL[6] = "\699_2_"
$JL[7] = "\683_2_"
$JL[8] = "\697_2_"
$JL[9] = "\800_2_"

;_ArrayDisplay($JL)

For $EL = 1 to 5

	Select

		Case $EL = 1

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[0] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[0] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[0] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[0] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[0], "", 1)
			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[5] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[5] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[5] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[5] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[5], "", 1)

			Next
			$File = ""

		Case $EL = 2

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[1] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[1] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[1] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[1] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[1], "", 1)
			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[6] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[6] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[6] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[6] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[6], "", 1)

			Next
			$File = ""

		Case $EL = 3

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[2] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[2] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[2] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[2] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[2], "", 1)
			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[7] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[7] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[7] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[7] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[7], "", 1)

			Next
			$File = ""

		Case $EL = 4

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[3] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[3] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[3] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[3] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[3], "", 1)
			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[8] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[8] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[8] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[8] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[8], "", 1)

			Next
			$File = ""

		Case $EL = 5

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[4] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[4] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[4] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[4] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[4], "", 1)
			FileMove("C:\APPS\EVO\job\EL" & $EL & $JL[9] & $File & ".twl", "C:\APPS\EVO\JLbackup" & $JL[9] & $File & ".twl", 1 + 8)
			FileCopy("C:\APPS\EVO\JLbackup" & $JL[9] & $File & ".twl", "C:\APPS\EVO\archief" & $JL[9] & $File & "_" & $filedatum & ".twl", 1 + 8)
			FileSetTime("C:APPS\EVO\JLbackup" & $JL[9], "", 1)

			Next
			$File = ""

	EndSelect

Next

EndFunc
;================== END FUNCTION move Joblist =====================================================================================================


;================== START FUNCTION Reset Joblist ========================================================================================================
Func ResetJoblist()

dim $EL
dim $File = ""
; n = 1/5, ELn folders: EL1, EL2, etc
; there is a maximum of 5 joblists per EL folder
; joblist $JL[x] is attached to ELn folder

Global $JL[10]

; joblist files EVO Blue
$JL[0] = "\682_1_"
$JL[1] = "\699_1_"
$JL[2] = "\683_1_"
$JL[3] = "\697_1_"
$JL[4] = "\800_1_"


; joblist files EVO Green
$JL[5] = "\682_2_"
$JL[6] = "\699_2_"
$JL[7] = "\683_2_"
$JL[8] = "\697_2_"
$JL[9] = "\800_2_"

;_ArrayDisplay($JL)

For $EL = 1 to 5

	Select

		Case $EL = 1

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\JLbackup" & $JL[0] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[0] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[0] & $File & ".twl", "", 1)
			FileMove("C:\APPS\EVO\JLbackup" & $JL[5] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[5] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[5] & $File & ".twl", "", 1)

			Next
			$File = ""

		Case $EL = 2

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\JLbackup" & $JL[1] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[1] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[1] & $File & ".twl", "", 1)
			FileMove("C:\APPS\EVO\JLbackup" & $JL[6] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[6] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[6] & $File & ".twl", "", 1)

			Next
			$File = ""

		Case $EL = 3

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\JLbackup" & $JL[2] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[2] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[2] & $File & ".twl", "", 1)
			FileMove("C:\APPS\EVO\JLbackup" & $JL[7] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[7] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[7] & $File & ".twl", "", 1)

			Next
			$File = ""

		Case $EL = 4

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\JLbackup" & $JL[3] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[3] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[3] & $File & ".twl", "", 1)
			FileMove("C:\APPS\EVO\JLbackup" & $JL[8] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[8] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[8] & $File & ".twl", "", 1)

			Next
			$File = ""

		Case $EL = 5

			For $File = 1 to 5

			FileMove("C:\APPS\EVO\JLbackup" & $JL[4] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[4] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[4] & $File & ".twl", "", 1)
			FileMove("C:\APPS\EVO\JLbackup" & $JL[9] & $File & ".twl", "C:\APPS\EVO\job\EL" & $EL & $JL[9] & $File & ".twl", 0 + 8)
			FileSetTime("C:\APPS\EVO\job\EL" & $EL & $JL[9] & $File & ".twl", "", 1)

			Next
			$File = ""

	EndSelect

next

EndFunc
;================== END FUNCTION Reset Joblist ================================================================================================

;================== START FUNCTION Delete Joblist ========================================================================================================
Func DeleteJoblist()

dim $EL
dim $File = ""
; n = 1/5, ELn folders: EL1, EL2, etc
; there is a maximum of 5 joblists per EL folder
; joblist $JL[x] is attached to ELn folder

Global $JL[10]

; joblist files EVO Blue
$JL[0] = "\682_1_"
$JL[1] = "\699_1_"
$JL[2] = "\683_1_"
$JL[3] = "\697_1_"
$JL[4] = "\800_1_"


; joblist files EVO Green
$JL[5] = "\682_2_"
$JL[6] = "\699_2_"
$JL[7] = "\683_2_"
$JL[8] = "\697_2_"
$JL[9] = "\800_2_"

;_ArrayDisplay($JL)

For $EL = 1 to 5

	Select

		Case $EL = 1

			For $File = 1 to 5

			FileDelete("C:\apps\EVO\job\EL" & $EL)

			Next

		Case $EL = 2

			For $File = 1 to 5

			FileDelete("C:\apps\EVO\job\EL" & $EL)

			Next

		Case $EL = 3

			For $File = 1 to 5

			FileDelete("C:\apps\EVO\job\EL" & $EL ) ;& "*.twl")

			Next

		Case $EL = 4

			For $File = 1 to 5

			FileDelete("C:\apps\EVO\job\EL" & $EL ) ;& "*.twl")

			Next

		Case $EL = 5

			For $File = 1 to 5

			FileDelete("C:\apps\EVO\job\EL" & $EL ) ;& "*.twl")

			Next
		Endselect
Next

EndFunc
;================== END FUNCTION Delete Joblist ================================================================================================

;================== Start Main() ==============================================================================================================

; GUI Creation
GuiCreate("JL Backup", 200, 200)
GUISetBkColor(0xE0FFFF)
GUISetFont(9, 500, 2, 45)
GuiCreate("EVO Joblist Backup", 200, 150)
GuiSetIcon(@WindowsDir & "\explorer.exe", 0)

; Button advanced, menu
$MoveJoblist = GUICtrlCreateButton("Move Joblists",20,15,165,25)
$ResetJoblist = GUICtrlCreateButton("Reset all Joblists",20,45,165,25)
$ExitButton = GUICtrlCreateButton("Exit",20,75,165,25)
$DeleteButton = GUICtrlCreateButton("Delete Joblists",20,105,165,25)

; Close Group
GUICtrlCreateGroup("",-99,-99,1,1)


GuiSetState(@SW_SHOW)
;GUISetState()

; Continuous Loop to check for GUI Events
While 1
$guimsg = GUIGetMsg()
Select

	Case $guimsg = $GUI_EVENT_CLOSE
	ExitLoop
	Case $guimsg = $MoveJoblist
	MoveJoblist()
	Exit
	Case $guimsg = $ResetJoblist
	Resetjoblist()
	Exit
	Case $guimsg = $ExitButton
	Exit
	Case $guimsg = $DeleteButton
		Dim $iMsgBoxAnswer
		$iMsgBoxAnswer = MsgBox(52,"DELETE Joblists"," Weet u dit heel erg zeker?")
		Select
			Case $iMsgBoxAnswer = 6 ;Yes
			DeleteJoblist()
			Case $iMsgBoxAnswer = 7 ;No
			;Exit
			;	What do I put in here to keep the program running?
		EndSelect
EndSelect
WEnd

;================== END Main() ==============================================================================================================