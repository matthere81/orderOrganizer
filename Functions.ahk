
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

SaveToIni:
; Gui Submit, NoHide
; if (!cpq) || (!po)
; {
;     MsgBox, Please enter a quote and PO#.
;     return
; }

return

EditChanged:
    ; Get the values of the CPQ and PO fields
    GuiControlGet, CPQ, , cpq
    GuiControlGet, PO, , po

    ; Check if both fields are not empty
    if (cpq != "" and po != "")
    {
        ; Start or reset a timer that calls the Autosave function after a 2-second delay
        SetTimer, Autosave, -3000
    }
    else
    {
        ; Stop the timer
        SetTimer, Autosave, Off
    }

Return

; This subroutine is called when an Edit field gains focus
FieldFocus:
    ; Get the name of the control that gained focus
    GuiControlGet, focusedControl, Focus
    MsgBox, The focus has changed to %focusedControl%.
Return

WM_LBUTTONDOWN(wParam, lParam)
{
    ; X := lParam & 0xFFFF
    ; Y := lParam >> 16
    if A_GuiControl
        Ctrl := "`n(in control " . A_GuiControl . ")"
    ToolTip You left-clicked in Gui window %A_Gui% %Ctrl%
    ; Sleep 2000
    ; ToolTip
}

; This subroutine is called when the user inputs text in the Edit field
; EditChanged:
;     IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . ".ini"
;     ; The user is currently inputting text
;     ; allInput := ""
;     ;Save the current content of the Edit field to a temporary variable
;     for index, var in vars
;     {
;         ; MsgBox, % IniFilePath
;         GuiControlGet fieldValue,, %var%
;         ; If the content has changed, display a message box
;         if (fieldValue != oldValues[var])
;         {
;             ; MsgBox, A new field has been entered: %fieldValue%
;             ; Update the old value
;             oldValues[var] := fieldValue
;         }
;         ; Append the content of the control to the string
;         ; allInput .= fieldValue . "`n"
;         IniWrite %fieldValue%, %IniFilePath%, orderInfo, %field%
        
;     }
;     ; MsgBox % allInput
; Return

Autosave:  

    ; GuiControl, SubCommand, ControlID [, Value]
    ; Show a message in the status bar
    SB_SetText("Autosaving...",,0)
    ; MsgBox in autosave
    
    IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . ".ini " ;. customer . ".ini"
    ; IniFilePathWithSo := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . " SO# " . soNumber . ".ini"

    ; fileNames := ""
    ; today := A_Now
    ; FormatTime, today, %today%, yyyyMMdd

    ; Loop, %myIniPath%\*.*
    ; {
    ;     FileGetTime, creationTime, %A_LoopFileLongPath%, C
    ;     FormatTime, creationTime, %creationTime%, yyyyMMdd

    ;     If (creationTime = today)
    ;     {
    ;         fileNames .= A_LoopFileName . "`n"
    ;     }
    ; }

    ; MsgBox, % fileNames

    if FileExist(IniFilePath)
    {
        MsgBox, 4,, A file with this quote and PO# already exists. Would you like to overwrite it?
        IfMsgBox, No
            return
    }

    ; Gosub EditChanged

    ; if FileExist(IniFilePath) ;or FileExist(IniFilePathWithSo)
    ; {
    ;     MsgBox, 4,, A file with this quote and PO# already exists. Would you like to overwrite it?
    ;     IfMsgBox, No
    ;         return
    ; }
    ; ; Else
    ; ; {
    ; ;     IniFilePath := myinipath . "\Temp.ini"
    ; ; }
    ; ; Return
    ; ; Determine the file path based on whether 'cpq' and 'po' are entered
    ; if (cpq and po)
    ; {
    ;     ; MsgBox In cpq and PO true
    ;     IniFilePath := IniFilePathWithSo
    ;     ; If the temporary file exists, delete it
    ;     if FileExist(myinipath . "\Temp.ini")
    ;         FileDelete, %myinipath%\Temp.ini
    ; }
    ; else
    ; {
    ;     MsgBox In else
    ;     IniFilePath := myinipath . "\Temp.ini"
    ; }

    

    ; Save the current state of the GUI to an ini file
    Gui Submit, NoHide
  
    ; fileNames := []
    ; myFile := "PO 456 CPQ-123.ini"
    ; msgbox % IniFilePath

    for index, var in vars
    {
        GuiControlGet, fieldValue,, %var%
        IniWrite %fieldValue%, %IniFilePath%, orderInfo, %var%
    }

    ; MsgBox % allInput
    Sleep 500
    SB_SetText("",,0) ; , 1)
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

moveDatabase(myIniPath)
{
    sourceDir := "C:\Users\" . A_UserName . "\Order Organizer\Order Database\*.*" ; The source directory
    broadDir := "C:\Users\" . A_UserName . "\Order Organizer"

    FileMove, %sourceDir%, %myIniPath%, 1 ; The '1' option overwrites existing files

    if ErrorLevel
    {
        MsgBox, 16, , An error occurred while moving the files.
    }
    else
    {
        MsgBox, , , The files were moved successfully.
    }

    FileRemoveDir, %broadDir%, 1
}