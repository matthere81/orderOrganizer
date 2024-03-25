
; Create the GUI
; Gui New,, Search ; Create a new GUI window with the title 'Search'

; Gui Show, , Search ; Show the GUI

; return ; End of auto-execute section

; Define the 'Search' subroutine
Search:
Gui New,, Search ; Create a new GUI window with the title 'Search'
Gui, Submit, NoHide ; Update the variables with the current values of the controls

folder := "C:\Users\" . A_UserName . "\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB"

; If the folder does not exist - creat the folder
if !FileExist(folder) ; If the folder does not exist
{
    FileCreateDir, %folder% ; Create the folder
    return
}

Run %folder% ; Open the selected folder

; Loop, %folder%\*.* ; Loop through all files in the selected folder
; {
;     IfInString, A_LoopFileName, %SearchTerm% ; If the current file name contains the search term
;         MsgBox, %A_LoopFileName% ; Display a message box with the file name
; }

return