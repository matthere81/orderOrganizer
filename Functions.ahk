
readtheini:
    if (SelectedFile = "") {
        ; MsgBox, You clicked Cancel.
        return
    }

    for index, var in vars
    {
        IniRead myValue, %SelectedFile%, orderInfo, %var%,
        if (myValue = "ERROR")
            myValue := ""
        GuiControl,, %var%, %myValue%
    }

Return

CheckFocus:
    ; ; Save the current state of the GUI to an ini file
    Gui Submit, NoHide
    GuiControlGet focusedControl, FocusV
    
    IniFilePath := SetIniFilePath(myinipath, po, cpq, customer)

    MsgBox % IniFilePath
    Gosub Autosave
    ; ; Read the previously saved value of the focused control from the INI file
    ; IniRead prevValue, %IniFilePath%, orderInfo, %focusedControl%

    ; ; Loop through all files in myinipath
    ; Loop, Files, %myinipath%\*, F
    ; {
    ;     ; Check if the current file matches IniFilePath
    ;     if (A_LoopFileFullPath = IniFilePath)
    ;         {
    ;             ; The file was found
    ;             matchedFilePath := A_LoopFileFullPath
    ;             ; MsgBox % customer . " and PO is " . po
    ;             ; Continue checking for a match until the variables are entered
    ;             while (po != "") && (cpq != "") && (customer != "") ;&& (crd != "") && (poDate != "") && (sapDate != "")
    ;             {
    ;                 ; MsgBox, 2
    ;                 ; Check if all fields in vars are not blank
    ;                 allFieldsEntered := false
    ;                 for index, var in autosaveVars
    ;                 {
    ;                     ; MsgBox % var
    ;                     ; if (var = "")
    ;                     ; {
    ;                     ;     ; MsgBox % var
    ;                     ;     allFieldsEntered := false
    ;                     ;     break
    ;                     ; }
    ;                 }
        
    ;                 ; ; If all fields in vars are entered, exit the while loop
    ;                 ; if (allFieldsEntered)
    ;                 ; {
    ;                 ;     break
    ;                 ; }
        
    ;                 ; TODO: Add code here to update the variables
    ;             }
        
    ;             break
    ;         }
    ; }
    
    ; ; Get the current value of the focused control
    ; GuiControlGet currentValue,, %focusedControl%
    ; ; MsgBox, %focusedControl%: %currentValue%
    
    ; if (cpq != "") && (po != "") && (focusedControl != "po") && (focusedControl != "cpq")  && (currentValue != prevValue)
    ; {
    ;     ; Show a message in the status bar
    ;     SB_SetText("",,0)
        ; Gosub Autosave
    ; }
Return

SaveToIni:
    Gui Submit, NoHide
    IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . ".ini " ;. customer . ".ini"
    for index, var in vars
    {
        GuiControlGet, fieldValue,, %var%
        msgbox % var
        IniWrite %fieldValue%, %IniFilePath%, orderInfo, %var%
    }
    MsgBox, Saved
return

Autosave:  
    ; Show a message in the status bar
    ; SB_SetText("Autosaving...",,0)
    
    ; ; Set the path for the ini file
    ; IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . ".ini " ;. customer . ".ini"
    
    ; ; Save the current state of the GUI to an ini file
    ; Gui Submit, NoHide
    
    ; allFields := "" ; Initialize an empty string to collect all fields and values

    for index, var in vars
    {
        GuiControlGet, fieldValue,, %var%
        IniWrite %fieldValue%, %TempIniFilePath%, orderInfo, %var%
        
        ; allFields .= "Field: " . var . ", Value: " . fieldValue . "`n" ; Add the field and value to the string
    }

    ; MsgBox % allFields ; Display all fields and values
    
    ; Sleep 500
    ; SB_SetText("",,0)
Return

SetIniFilePath(myinipath, po, cpq, customer)
{
    ; Set the path for the ini file
    if (po != "") && (cpq != "") && (customer != "")
    {
        IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . ".ini"
    }
    else if (po != "") && (cpq != "")
    {
        IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . ".ini"
    }
    else
    {
        IniFilePath := myinipath . "\Temp.ini"
    }
    return IniFilePath
}

InitialState:
; Save the initial state of the controls after the data is loaded
Gui Submit, NoHide
initialState := {}
initialStateString := ""
for index, var in vars
{
    GuiControlGet, fieldValue,, %var%
    initialState[var] := fieldValue
    initialStateString .= var . ": " . fieldValue . "`n"
}
MsgBox, InitialState:`n%initialStateString%
Return

OpenFileFromMenu:
    FileSelectFile, SelectedFile,r,%myinipath%, Open a file
    Gosub readtheini
Return

ShowChecklists:
    ; Get the current height of the GUI
    Gui, +LastFound
    WinGetPos, , , , currentHeight

    if (currentHeight > guiHeight and currentHeight < guiChecklistHeight)
    {
        Gui Show, y150 w%guiWidth% h%guiChecklistHeight%, Order Organizer
    }
    else if (currentHeight > guiChecklistHeight)
    {
        Gui Show, y300 w%guiWidth% h%guiHeight%, Order Organizer
    }
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
    GuiControl,, % endUse
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


AddSection(section, checklist, resetXCoord = false)
{
    ; Initialize rightmost edge of the furthest section
    static rightmostEdge := 10

    ; Reset the x-coordinate if requested
    if (A_Index = 1)
        rightmostEdge := 10

    ; Calculate y-coordinate based on section index
    yCoord := 575 ; Adjust multiplier as needed for spacing
    xCoord := rightmostEdge ;10 + (A_Index - 1) * 270 ; Adjust multiplier as needed for spacing
    checkboxYCoord := 600 + (A_Index - 1) ;* 20 ; Adjust multiplier as needed for spacing
    index := A_Index

    ; Calculate the width of the longest string in the checklist
    maxWidth := 25
    Loop, % checklist.MaxIndex()
    {
        item := checklist[A_Index]
        if (StrLen(item) > maxWidth)
            maxWidth := StrLen(item)
            ; MsgBox % maxWidth
    }

    ; Convert the width from characters to pixels
    ; This is a rough approximation, adjust the multiplier as needed
    maxWidth := maxWidth * 8

    ; Add a group box for the section with the calculated width
    Gui Add, GroupBox, x%xCoord% y%yCoord% h175 w%maxWidth%, % section ;[A_Index]

    ; Update the rightmost edge of the furthest section
    rightmostEdge := xCoord + maxWidth + 5 ; Add some padding

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
        checkboxYCoord := checkboxYCoord + 25 ; Adjust as needed for spacing
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

backupDatabase(myinipath)
{
    sourceDir := % myinipath . "\*.*" ; The source directory
    backupDir := "C:\Users\" . A_UserName . "\OneDrive - Thermo Fisher Scientific\Documents\Order Organizer - Backup", 1 ; The '1' option overwrites existing files
    ; FileCopy, Source, Dest [, Flag (1 = overwrite)]
    if !FileExist(backupDir)
    {
        FileCreateDir, %backupDir%
    }
    FileCopy, %sourceDir%, %backupDir%, 1
}