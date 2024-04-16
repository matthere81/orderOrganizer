
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
    ; MsgBox, % section . " " . checklist
    ; Calculate y-coordinate based on section index
    yCoord := 575 ; Adjust multiplier as needed for spacing
    xCoord := 10 + (A_Index - 1) * 300 ; Adjust multiplier as needed for spacing
    checkboxYCoord := 600 ;+ (A_Index - 1) * 20 ; Adjust multiplier as needed for spacing
    index := A_Index

    ; Add a group box for the section
    Gui Add, GroupBox, x%xCoord% y%yCoord% h175 w280, % section ;[A_Index]

    ; Print the value of checklist[A_Index].MaxIndex() for debugging
    ; MsgBox, % "checklist[A_Index].MaxIndex(): " . checklist[A_Index].MaxIndex() 

   ; Create a global array to store the checkbox states
    global checkboxStates := []

    ; Loop over each item in the checklist
    Loop, % checklist.MaxIndex()
        {
            ; Calculate x-coordinate for the checkbox based on item index
            checkboxXCoord := % xCoord + 10 ;+ (A_Index - 1) ; * 15 ; Adjust multiplier as needed for spacing
            
            ; Get the checklist item
            item := checklist[A_Index]
        
            ; Add the checkbox state to the array
            checkboxStates.Push(false)
        
            ; Create a checkbox with the item text and a g-label to update the checkbox state
            Gui Add, Checkbox, x%checkboxXCoord% y%checkboxYCoord% v%varName% gUpdateCheckboxState, %item%
        
            ; Increase the y-coordinate for the next checkbox
            checkboxYCoord := checkboxYCoord + 20 ; Adjust as needed for spacing
        }
}
return

; Define the g-label function to update the checkbox state
UpdateCheckboxState:
    ; Get the index of the checkbox that was checked or unchecked
    ControlGet checkboxIndex, Choice, , % A_GuiControl

    ; Update the checkbox state in the array
    checkboxStates[checkboxIndex] := !checkboxStates[checkboxIndex]
return

sanitizeVarName(item) {
    ; Create an array of characters to replace
    charsToReplace := [" ", "&", "(", ")", "-", "/", "?"]

    ; Create a sanitized version of the item to use as the variable name
    varName := item

    ; Loop over each character in the array
    Loop, % charsToReplace.MaxIndex()
    {
        ; Replace the current character with an underscore
        varName := StrReplace(varName, charsToReplace[A_Index], "_")
    }

    return varName
}