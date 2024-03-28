
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
    ; SelectedFile := myinipath . "\" . SelectedFile
    ; MsgBox % SelectedFile

    for index, var in vars
    {
        IniRead, value, %SelectedFile%, orderInfo, %var%
        ; MsgBox % value
        ; if (field == "sot" && value == "ERROR")
            ; value := 
        GuiControl,, %value%, % A_Index
        ; GuiControl, SubCommand, ControlID , Value
        
        ; MsgBox % value . " " . var . " " . index . " " . A_Index
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
    SelectedRow := LV_GetNext() ; Get the selected row number
    LV_GetText(SelectedFile, SelectedRow, 1) ; Get the text of the selected item from the first column
    SelectedFile := myinipath . "\" . SelectedFile ;. ".ini" ; Add the '.ini' extension to the selected file
    Gosub readtheini
    Gui, %MyGui%:Destroy ; Destroy the new GUI
return