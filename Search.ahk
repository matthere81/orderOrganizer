Search:
Gui Submit, NoHide ; Update the variables with the current values of the controls
SearchTerm := Trim(SearchTerm) ; Remove trailing whitespace from SearchTerm

folder := "C:\Users\" . A_UserName . "\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB"

; If the folder does not exist - create the folder
if !FileExist(folder) ; If the folder does not exist
{
    FileCreateDir, %folder% ; Create the folder
    return
}

Loop, %folder%\*.* ; Loop through all files in the selected folder
{
    IfInString, A_LoopFileName, %SearchTerm% ; If the current file name contains the search term
        MsgBox, %A_LoopFileName% ; Display a message box with the file name
}

return