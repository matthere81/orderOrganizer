Search:
Gui Submit, NoHide ; Update the variables with the current values of the controls
SearchTerm := Trim(SearchTerm) ; Remove trailing whitespace from SearchTerm

if (SearchTerm = "") ; If there's no matching files
{
    MsgBox, No matching files were found. ; Display a message box
    Return
}

matchingFiles := [] ; Create an empty array to store matching files

Loop, %myinipath%\\*.* ; Loop through all files in the selected folder
{
    IfInString, A_LoopFileName, %SearchTerm% ; If the current file name contains the search term
    {
        matchingFiles.Push(A_LoopFileName) ; Add the file name to the list of matching files
    }
}

    if (matchingFiles.Length() = 1) ; If there's only one matching file
    {
        SelectedFile := matchingFiles[1] ; Set the 'file' variable to the name of the file
        Gosub readtheini
    }
    else if (matchingFiles.Length() > 1) ; If there's more than one matching file
    {
        Gui New, +HwndMyGui ; Create a new GUI and store its handle in the variable 'MyGui'
        Gui Font, S9, Segoe UI ;Semibold
        Gui Color, 79b8d1
        Gui Add, ListView, vSelectedFile gFileSelected w475, Search Results ; Add a ListBox control
        Loop % matchingFiles.Length() ; Loop over the 'matchingFiles' array
        {
            LV_Add("", matchingFiles[A_Index]) ; Add the file name to the ListView
        }
        Gui Show, w500 h200 ; Show the GUI
        ; Hotkey IfWinActive, ahk_id %MyGui% ; Set the hotkey to be active only when the new GUI is open
        ; Hotkey Escape, GuiClose, On ; Set the Escape key to trigger the GuiClose label
        ; Hotkey Enter, FileSelected, On ; Set the Enter key to trigger the FileSelected label
    }
return

FileSelected:
    Gui Submit, ;NoHide ; Get the selected file
    LV_GetText(SelectedFile, 1) ; Get the text of the selected item
    Gosub readtheini
    ; WinActivate, ahk_id %MyGui% ; Activate the new GUI again
return

; GuiClose:
;     ; Hotkey, Escape, Off ; Turn off the Escape hotkey
;     Hotkey, Enter, Off ; Turn off the Enter hotkey
;     Gui, %MyGui%:Destroy ; Destroy the new GUI
; return