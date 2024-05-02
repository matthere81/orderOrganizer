
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

        if (var = "endUse" or var = "notes")
        {
            ; Handle endUse and notes differently
            StringReplace, myValue, myValue, ``n, `n, All
            StringReplace, myValue, myValue, ``r, `r, All
        }

        GuiControl,, %var%, %myValue%
    }

    ; Read the checklist variables
    for outerIndex, checklists in allChecklists
    {
        for midIndex, checklistArray in checklists
        {
            for innerIndex, checklist in checklistArray
            {
                IniRead myValue, %SelectedFile%, orderInfo, %checklist%,
                ; MsgBox, % myValue
                if (myValue != "ERROR")
                GuiControl,, %checklist%, %myValue%
            }
        }
    }
Return

TrimAndGetValue(controlName)
{
    GuiControlGet value,, %controlName%
    value := Trim(value)
    return value
}

GuiState:
    Gui Submit, NoHide
    GuiControlGet focusedControl, FocusV
return

SaveToIni:
    Gosub GuiState
    IniFilePath := SetIniFilePath(myinipath, po, cpq, customer, vars)
    save(vars, allChecklists, IniFilePath)
    checkIfOrderFolderExists(myOrderDocs, po, cpq, customer)
return

save(vars, allChecklists, IniFilePath)
{
    for index, var in vars
    {
        GuiControlGet, fieldValue,, %var%
        fieldValue := Trim(fieldValue) ; Trim whitespace from fieldValue

        if (var = "endUse" or var = "notes")
        {
            ; Handle endUse and notes differently
            StringReplace, fieldValue, fieldValue, `n, ``n, All
            StringReplace, fieldValue, fieldValue, `r, ``r, All
        }

        IniWrite %fieldValue%, %IniFilePath%, orderInfo, %var%
    }

    ; Save the checklist variables
    for outerIndex, checklists in allChecklists
    {
        for midIndex, checklistArray in checklists
        {
            for innerIndex, checklist in checklistArray
            {
                GuiControlGet, fieldValue,, %checklist%
                IniWrite %fieldValue%, %IniFilePath%, orderInfo, %checklist%
            }
        }
    }

    saveStatusBar()
}

saveStatusBar()
{
    ; Update the status bar
    GuiControl,, MyStatusBar, Saving...
    
    ; Pause the script for two seconds
    Sleep 1000
    
    ; Clear the status bar
    GuiControl,, MyStatusBar, 
    Return
}

SetIniFilePath(myinipath, po, cpq, customer, vars)
{
    ; Check if po or cpq are empty
    if (po = "" or cpq = "")
    {
        ; Display a MsgBox and return from the function
        MsgBox, Please enter a quote# and PO# to enable saving.
        return
    }

    IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq
    if (customer != "")
    {
        IniFilePath .= " " . customer
    }
    IniFilePath .= ".ini"

    if customer = ""
    {
        compareFilenames(IniFilePath, myinipath)
    }
    return IniFilePath
}

compareFilenames(IniFilePath, myinipath)
{
    ; Extract the filename from IniFilePath
    SplitPath IniFilePath, targetFileName
    
    Loop Files, %myinipath%\*.*
    {
        SplitPath A_LoopFileLongPath, currentFileName
        if (currentFileName = targetFileName)
        {
            MsgBox, 4, Match found, % "The following match was found:`n`n" . currentFileName 
                . "`n`nDo you want to keep your current file and delete the old file?"
                IfMsgBox, Yes
                {
                    ; User chose to keep going with the new file and delete the old file
                    FileDelete, %A_LoopFileLongPath%
                    return targetFileName
                }
                else
                {
                    ; User chose to run the gClearFields subroutine
                    Gosub ClearFields
                    return
                }
        }
    }
    return
}

checkIfOrderFolderExists(myOrderDocs, po, cpq, customer)
{
    MsgBox, % po . " is PO`n" . cpq  . " is CPQ`n" . customer . " is customer`n" . myOrderDocs . " is myOrderDocs`n" . droppedFile
    if (po = "" or cpq = "")
    {
        MsgBox, Please add PO & CPQ/Quote#.
        return
    }
    ; Create the order folder path
    folderPath := myOrderDocs . "\PO " . po . " " . customer . " CPQ-" . cpq
    ; Check if the order folder exists
    if !FileExist(folderPath)
    {
        ; Create the order folder
        FileCreateDir, %folderPath%
        run % folderPath
    }
    Return folderPath
}

GuiDropFiles(GuiHwnd, FileArray, CtrlHwnd, X, Y) ;, po, cpq, customer, myOrderDocs)
{
    guiDropTemp := "C:\Users\" . A_UserName . "\Order Organizer\Temp"
    if !FileExist(guiDropTemp) {
        FileCreateDir, %guiDropTemp%
    }

    for i, file in FileArray
        FileCopy, %file%, %guiDropTemp%
    return
}

ExtractAllAttachmentsFromCurrentEmail(PathToSaveTo)
{
    guiDropTemp := "C:\Users\" . A_UserName . "\Order Organizer\Temp"
    if !FileExist(guiDropTemp) {
        FileCreateDir, %guiDropTemp%
    }

    try
    {
        ; Try to get a COM object for the running instance of Outlook
        outlook := ComObjActive("Outlook.Application")
        ; MsgBox, Outlook is open.
    }
    catch
    {
        ; If an exception is thrown, Outlook is not open
        MsgBox, Please open Outlook and try again.
        return
    }

    ; Create a COM object for the running instance of Outlook
    outlook := ComObjActive("Outlook.Application")
    ; Get the currently selected item
    email := outlook.ActiveExplorer.Selection.Item(1)

    ; Check if the item is a mail item
    if (email.Class == 43) ; 43 is the class type for a mail item
    {
        ; Get the subject of the email
        subject := email.Subject

        ; Define an array of strings to look for in the subject
        searchStrings := ["cpq", "purchase order number", "purchase order #", "purchase order","po number", "po#", "po #", "po"]

        findInfoFromSubject(subject, searchStrings, potentialPo)

        ; MsgBox, % potentialPo

        ; Check if the mail item has attachments
        if (email.Attachments.Count > 0)
        {
            saveAttachments(email, guiDropTemp)
            processFiles(guiDropTemp, searchStrings, potentialPo)
            ; MsgBox % "PO is " . po . "`n" . "CPQ is " . cpq
            ; global cpq := cpq
            GuiControl,, po, %po%
            GuiControl,, cpq, %cpq%
            if (cpq != "")
            {
                quoteId := cpq
                getQuoteInfo(quoteID, contactName, contactEmail, contactPhone, customerName, quoteOwner, creatorManager, totalNetAmount, totalFreight, surcharge, totalTax, quoteTotal, soldToID, paymentTerms, opportunity)

            }
        }
    }
    else
    {
        MsgBox, No email is currently selected or the selected item is not an email.
    }
}

findInfoFromSubject(ByRef subject, searchStrings, ByRef potentialPo)
{
    ; MsgBox % subject . " is the subject." . "`n" . searchStrings . " is the searchStrings."
    ; Loop through the array and check if the subject contains any of the strings
    for index, searchString in searchStrings
    {
        ; MsgBox % subject . " is the subject - in the for loop."
        ; Define a regular expression that matches the searchString followed by any characters
        regex := "(?i)" searchString "(?:\s*-\s*|\s*)(\S+)" 

        if (InStr(subject, searchString))
        {
            if (searchString = "cpq")
            {
                regex := "(?i)" searchString "\s*-\s*(\d{8})"
                ; Extract the CPQ number from the subject
                extractNumbersFromSubject(subject, regex, searchString, potentialPo)
            }
            else
            {
                ; MsgBox % subject . " is the subject - in the ELSE."
                ; Extract the PO number from the subject
                extractNumbersFromSubject(subject, regex, searchString, potentialPo)
                break
            }
        }
    }
}

extractNumbersFromSubject(ByRef subject, regex, searchString, ByRef potentialPo)
{
    ; MsgBox % subject . " is the subject. In the extractNumbersFromSubject function."
    RegExMatch(subject, regex, match)
    ; Extract the number directly after the match
    Trim(match1)
    if (searchString = "cpq")
    {
        potentialPo := match1
        potentialQuote := potentialPo
        if (potentialPo = match1)
        {
            cpq := potentialQuote
            ; MsgBox % cpq
            return cpq
        }
        ; MsgBox, % potentialQuote
        return potentialQuote
    }
    Else
    {
        if (potentialPo = match1)
        {
            po := potentialPo
            ; MsgBox % po
            return po
        }
        
        potentialPo := match1
        return potentialPo
    }
}


saveAttachments(email, guiDropTemp)
{
    ; Loop through the attachments
    for thisattachment in email.Attachments
    {
        ; Save the attachment to the specified path
        attachmentPath := guiDropTemp . "\" . thisattachment.DisplayName
        thisattachment.SaveAsFile(attachmentPath)
    }
}

processFiles(guiDropTemp, searchStrings, potentialPo)
{
    ; Loop through the files in the attachment path
    Loop, Files, %guiDropTemp%\*.*
    {
        processFile(A_LoopFileLongPath, searchStrings, potentialPo)
    }
}

processFile(filePath, searchStrings, potentialPo)
{
    ; Get the base name and extension of the file
    SplitPath, filePath, name, dir, ext, name_no_ext, drive

    ; Check if the file extension is "png" or "jpg"
    if (ext = "png" || ext = "jpg" || ext = "gif")
    {
        ; Delete the file
        FileDelete, %filePath%
    }

    ; Check if the file name contains "po" or "cpq"
    if (InStr(name, "po") || InStr(name, "cpq")) ; || InStr(name, "PO") || InStr(name, "CPQ"))
    {   
        convertPdfToTextAndProcess(filePath, searchStrings, potentialPo)
    }
}

convertPdfToTextAndProcess(filePath, searchStrings, potentialPo)
{
    pdftotextPath:= % A_WorkingDir . "\pdftotext.exe"

    ; Convert the PDF to a text file using pdftotext
    RunWait, %pdftotextPath% -layout "%filePath%" "%filePath%.txt"
    
    Loop, Read, %filePath%.txt
    {
        ; Append lines that contain "po" or "cpq" to output.txt
        ; if (InStr(A_LoopReadLine, "po number") || InStr(A_LoopReadLine, "cpq") || InStr(A_LoopReadLine, "po #"))
        ; {
            ; line := A_LoopReadLine
            ; MsgBox % line
            subject := A_LoopReadLine

            findInfoFromSubject(ByRef subject, searchStrings, ByRef potentialPo)
        ; }
    }
}

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

    ; Clear the checkboxes
    for outerIndex, checklists in allChecklists
    {
        for midIndex, checklistArray in checklists
        {
            for innerIndex, checklist in checklistArray
            {
                GuiControl,, %checklist%, 0 ; Uncheck the checkbox
            }
        }
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

ExtractAllAttachmentsFromCurrentEmail:
    ExtractAllAttachmentsFromCurrentEmail(myOrderDocs)
Return

FileSelected:
    SelectedRow := LV_GetNext() ; Get the selected row number
    LV_GetText(ListView, SelectedRow, 1) ; Get the text of the selected item from the first column
    SelectedFile := myinipath . "\" . ListView ;. ".ini" ; Add the '.ini' extension to the selected file
    Gui, %MyGui%:Destroy ; Destroy the new GUI
    Gosub ClearFields
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