JLBackup, for Tecan EVOware & LIMS joblist management

To be compiled with AutoIt version 3.3.14.2
Please check if all included libraries are installled on your machine if compiling fails.
_________________________________________________________________________________________
prerequisites:

joblists as *.twl files in ELx folders
joblist filename: worklistnumber_robot_number.twl, for example 682_1_1.twl

C:\apps\evo\job\ELx (x=1,2,3, etc)
C:\apps\evo\JLBackup\
C:\apps\evo\JLBackup\JLbackup.ini
C:\ProgramData\Tecan\Evoware\AuditTrail\LOG\*.Log
_________________________________________________________________________________________
Button functionality:

"Set Joblists 1 to 5" : renames the Joblists send by the LIMS per folder from 1 to max 5

"Move joblists", moves all joblists from C:\apps\evo\job\ELx to C:\apps\evo\job\JLBackup

"Reset all joblists" moves from C:\apps\evo\job\JLBackup to C:\apps\evo\job\ELx 

"Delete Joblists", deletes all joblists from C:\apps\evo\job\ELx 

"Exit", exits program

right hand buttons open the filefolders of the same name in C:\apps\evo\<buttonname>
__________________________________________________________________________________________

ini file is made upon running JL Backup if non-existing, i.e. set to default upon deletion.
ini file defines LIMS worklistnumbers per ELx joblistfolder fro JLBackup.

ELx = \worklistnumber_robotnumber_ (EL1 = \682_1_)
_________________________________________________________________________________________
Future addition: filelocations added to ini file to allow for customized use.
_________________________________________________________________________________________
Dion Methorst, Amsterdam, oct 2015.
