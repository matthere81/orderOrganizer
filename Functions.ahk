
readtheini:
    ; Gui, Submit, NoHide
    ; if (cpq) && (po)
    ;     gosub, SaveToIniNoGui
    ; FileSelectFile, SelectedFile,r,%myinipath%, Open a file
    if (SelectedFile = "") {
        ; MsgBox, You clicked Cancel.
        return
    }
    ; MsgBox % SelectedFile
    SelectedFile := myinipath . "\" . SelectedFile
    ; MsgBox % SelectedFile

    for index, var in vars
    {
        IniRead, value, %SelectedFile%, orderInfo, %var%
        ; if (field == "sot" && value == "ERROR")
            ; value := 
        GuiControl,, %A_Index%, % value
        ; GuiControl, SubCommand, ControlID , Value
        
        ; MsgBox % value
    }
Return

; IniRead, OutputVar, Filename, Section, Key , Default
; IniRead, OutputVarSection, Filename, Section
; IniRead, OutputVarSectionNames, Filename

OpenFileFromMenu:
    FileSelectFile, SelectedFile,r,%myinipath%, Open a file
    Gosub readtheini
Return

ShowChecklists:
    MsgBox placeholder for checklist toggle setting
Return

Restart:
    Reload ; Restart the script
return

SetSearchAsDefault:
    KeyWait enter, d
    GuiControl Focus, Search ; Set the focus to the Search button
return

FileSelected:
    Gui Submit, ;NoHide ; Get the selected file
    LV_GetText(SelectedFile, 1) ; Get the text of the selected item
    Gosub readtheini
return