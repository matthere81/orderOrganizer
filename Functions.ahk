
readtheini:

    if (SelectedFile = "") {
        ; MsgBox, You clicked Cancel.
        return
    }

    for index, var in vars
    {
        IniRead myValue, %SelectedFile%, orderInfo, %var%,
        GuiControl,, %var%, %myValue%
    }
Return

SaveToIni:
Gui, Submit, NoHide
; if (!cpq) || (!po)
; {
;     MsgBox, Please enter a quote and PO#.
;     return
; }
IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . ".ini"
IniFilePathWithSo := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . " SO# " . soNumber . ".ini"

; IniWrite, %sot%, %IniFilePath%, orderInfo, sot
; IniWrite, %customer%, %IniFilePath%, orderInfo, customer

; if FileExist(IniFilePath) && (soNumber)
; {
;     ; gosub, WriteIniVariables
;     ; FileMove, %IniFilePath%, %IniFilePathWithSo% , 1
;     ; IniFilePath = %IniFilePathWithSO% 
;     ; Gosub, SaveBar
;     ; gosub, CheckIfFolderExists
;     return  
; }
; else if FileExist(IniFilePathWithSo)
; {
;     ; IniFilePath = %IniFilePathWithSO% 
;     ; Gosub, SaveBar
;     ; gosub, WriteIniVariables
;     return
; }
; else if FileExist(IniFilePath) && (!soNumber)
; {
;     ; gosub, WriteIniVariables
;     ; Gosub, SaveBar
;     ; gosub, CheckIfFolderExists
;     return
; }
; else if !FileExist(IniFilePath) && !FileExist(IniFilePathWithSo)
; {
;     ; if(soNumber)
;     ; {
;     ;     IniFilePath = %IniFilePathWithSo%
;     ; } else {
;     ;     IniFilePath = %IniFilePath%
;     ; }
;     ; gosub, WriteIniVariables
;     ; Gosub, SaveBar
;     ; gosub, CheckIfFolderExists
;     return
; }
return

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

DeleteFile:
    ; FileSelectFile, SelectedFile,r,%myinipath%, Open a file
    ; FileDelete, %SelectedFile%
    ; MsgBox, File deleted
return

ArchiveFiles:
; Get the current time
CurrentTime := A_Now

; Loop over all files in the myinipath directory
Loop, Files, %myinipath%\*
{
    ; Get the creation time of the file
    FileGetTime, CreationTime, %A_LoopFileFullPath%, C

    ; Calculate the difference between the current time and the creation time in years
    YearsOld := (CurrentTime - CreationTime) // 31536000

    ; If the file is more than three years old
    if (YearsOld >= 3)
    {
        ; Get the year the file was created
        YearCreated := SubStr(CreationTime, 1, 4)

        ; If the folder for the year doesn't exist, create it
        if !FileExist(myinipath . "\" . YearCreated)
        {
            FileCreateDir, %myinipath%\%YearCreated%
        }

        ; Move the file to the folder for the year
        FileMove, %A_LoopFileFullPath%, %myinipath%\%YearCreated%
    }
}
return

SetSearchAsDefault:
    KeyWait enter, d
    GuiControl Focus, Search ; Set the focus to the Search button
return

FileSelected:
    SelectedRow := LV_GetNext() ; Get the selected row number
    LV_GetText(ListView, SelectedRow, 1) ; Get the text of the selected item from the first column
    SelectedFile := myinipath . "\" . ListView ;. ".ini" ; Add the '.ini' extension to the selected file
    Gui, %MyGui%:Destroy ; Destroy the new GUI
    Gosub readtheini
return