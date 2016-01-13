JLBackup v2, for Tecan EVOware & LIMS joblist management

See remarks section before script in JLBackup.au3
________________________________________________________________________________________________
To be compiled with AutoIt version 3.3.14.2
Please check if all included libraries are installled on your machine if compiling fails.
________________________________________________________________________________________________
prerequisites:

joblists as *.twl files in ELx folders
joblist filename: worklistnumber_robot_number.twl, for example 682_1_1.twl

C:\apps\evo\job\ELx (x=1,2,3, etc)
________________________________________________________________________________________________
Button functionality:

1st tab:
"Set Joblists 1 to 5" : renames the Joblists send by the LIMS per folder from 1 to max 5
"Delete Joblists", deletes all joblists from C:\apps\evo\job\ELx 
"Exit", exits program

2nd tab:
5 buttons on the left side and 5 buttons on the right side

These buttons can be assigned in the JLbackup.ini file
_____________________________________________________________________________________________
JLBackup.ini file location is now the same as the script location
JlBackup.ini default is created after deletion of the ini file and upon running JLBackup.exe

[Blue] & [Green] sections define LIMS worklistnumbers per ELx joblistfolder fro JLBackup.

folder ELx = \worklistnumber_robotnumber_ (EL1 = \682_1_)

[ShortcutsL] section for ButtonL1 to ButtonL5 and [ShortcutsR] section for ButtonR1 to ButtonR5,
followed by five keys for the button names with accompanying values for the paths to open

ButtonName=PathToOpen i.e. JOB=C:\APPS\EVO\JOB
________________________________________________________________________________________________
Dion Methorst, Sanquin IPB/CAP, Amsterdam, jan 2016.
