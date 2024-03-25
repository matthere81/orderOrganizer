
readtheini:
Gui, Submit, NoHide
; if (cpq) && (po)
;     gosub, SaveToIniNoGui
FileSelectFile, SelectedFile,r,%myinipath%, Open a file
; if (ErrorLevel)
;     {
;         gosub, restartScript
;         return
;     }

for index, var in vars
{
    IniRead, value, %SelectedFile%, orderInfo, %var%
    ; if (field == "sot" && value == "ERROR")
        ; value := 
    GuiControl,, %A_Index%, % value
    ; GuiControl, SubCommand, ControlID , Value
    
    ; MsgBox % value
}

; IniRead, OutputVar, Filename, Section, Key , Default
; IniRead, OutputVarSection, Filename, Section
; IniRead, OutputVarSectionNames, Filename