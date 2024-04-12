
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

; ; Write the values to the INI file
; for index, key in vars
; {
;     IniWrite, %values[key]%, %iniFilePath%, EditFields, %key%
; }

; MsgBox % IniFilePath . "`n" . IniFilePathWithSo . "`n" . soNumber

; if FileExist(IniFilePath) && (soNumber)
; {
;     MsgBox, first if
;     ; gosub, WriteIniVariables
;     ; FileMove, %IniFilePath%, %IniFilePathWithSo% , 1
;     ; IniFilePath = %IniFilePathWithSO% 
;     ; Gosub, SaveBar
;     ; gosub, CheckIfFolderExists
;     return  
; }
; else 
if FileExist(IniFilePathWithSo)
{
    MsgBox, second if
    ; IniFilePath = %IniFilePathWithSO% 
    ; Gosub, SaveBar
    ; gosub, WriteIniVariables
    return
}
; else if FileExist(IniFilePath) && (!soNumber)
; {
;     MsgBox, third if
;     ; gosub, WriteIniVariables
;     ; Gosub, SaveBar
;     ; gosub, CheckIfFolderExists
;     return
; }
; else if !FileExist(IniFilePath) && !FileExist(IniFilePathWithSo)
; {
;     msgbox, fourth if
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

ClearFields:
for index, field in vars
{
    GuiControl,, %field%, 
}
GuiControl,, SearchTerm,
return

ButtonDefault:
    GuiControlGet,FocusControl,Focus
    If (FocusControl!="SysListView321")
        Return
    Gosub FileSelected
Return

FileSelected:
    SelectedRow := LV_GetNext() ; Get the selected row number
    LV_GetText(ListView, SelectedRow, 1) ; Get the text of the selected item from the first column
    SelectedFile := myinipath . "\" . ListView ;. ".ini" ; Add the '.ini' extension to the selected file
    Gui, %MyGui%:Destroy ; Destroy the new GUI
    Gosub readtheini
return


;|------------------------------------------|
;|                                          |
;|          Add Checklist & Section         |
;|                                          |
;|------------------------------------------|

AddSection(section, checklist)
{
    MsgBox, % section . " " . checklist
    ; Calculate y-coordinate based on section index
    yCoord := 575 ; Adjust multiplier as needed for spacing
    xCoord := 10 + (A_Index - 1) * 300 ; Adjust multiplier as needed for spacing

    ; Add a group box for the section
    Gui Add, GroupBox, x%xCoord% y%yCoord% h175 w280, % section ;[A_Index]

    checkboxYCoord := 595 ;+ (A_Index - 1) * 20 ; Adjust multiplier as needed for spacing
    index := A_Index

    ; Print the value of checklist[A_Index].MaxIndex() for debugging
    ; MsgBox, % "checklist[A_Index].MaxIndex(): " . checklist[A_Index].MaxIndex() 

    ; Loop over each item in the checklist
    Loop, % checklist[A_Index].MaxIndex()
    {
        MsgBox, here
        ; Calculate x-coordinate for the checkbox based on item index
        checkboxXCoord := % xCoord + 10 ;+ (A_Index - 1) ; * 15 ; Adjust multiplier as needed for spacing
        
        ; Get the checklist item
        item := checklist[index,A_Index] ;,[A_Index[A_Index]] ;,%index%]
        ; MsgBox, % index . " " . A_Index . " " . item
        ; Create a sanitized version of the item to use as the variable name
        varName := StrReplace(item, " ", "_")
        varName := StrReplace(varName, "&", "_")
        varName := StrReplace(varName, "(", "_")
        varName := StrReplace(varName, ")", "_")
        varName := StrReplace(varName, "-", "_")
        varName := StrReplace(varName, "/", "_")
        varName := StrReplace(varName, "?", "_")

        ; ; If the item text is longer than 45 characters, split it into two lines
        ; if (StrLen(item) > 40)
        ; {
        ;     ; Find the position of the next whitespace character after the 45th character
        ;     pos := InStr(item, " ", false, 41)
        ;     ; If a whitespace character was found, split the item text at that position
        ;     if (pos > 0)
        ;     {
        ;         line1 := SubStr(item, 1, pos - 1)
        ;         line2 := SubStr(item, pos + 1)
        ;     }
        ;     else
        ;     {
        ;         ; If no whitespace character was found, split the item text at the 45th character
        ;         line1 := SubStr(item, 1, 45)
        ;         line2 := SubStr(item, 46)
        ;     }
        
        ;     Gui Add, Checkbox, x%checkboxXCoord% y%checkboxYCoord% v%varName%, %line1%`n%line2%
        ;     ; Increase the y-coordinate for the next checkbox by an additional amount to account for the extra line
        ;     checkboxYCoord := checkboxYCoord + 30 ; Adjust as needed for spacing
        ; }
        ; else
        ; {
            Gui Add, Checkbox, x%checkboxXCoord% y%checkboxYCoord% v%varName%, %item%
            ; Increase the y-coordinate for the next checkbox
            checkboxYCoord := checkboxYCoord + 20 ; Adjust as needed for spacing
        ; }
    }
}
return