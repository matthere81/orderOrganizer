#SingleInstance Force

; Create a GUI
Gui, Add, Edit, vSearchEdit w380 -VScroll gSearchFiles
Gui, Add, ListView, r20 w400 vMyListView Grid, File Name

; Specify the directory to search
global directoryPath := "C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB" ; Replace with the directory path to search

; Show the GUI
Gui, Show, w420 h300, File Search
return

; Function to search for files within the specified directory
SearchFiles(directoryPath, searchTerm) {
    LV_Delete() ; Clear the ListView
    Loop, Files, %directoryPath%\*.*
    {
        filePath := A_LoopFileFullPath
        fileName := A_LoopFileName

        ; Check if the file matches the search term
        if (InStr(fileName, searchTerm))
        {
            LV_Add("", fileName) ; Add the filename to the ListView
        }
    }
}

; Function to handle the search text change
SearchFiles:
    Gui, Submit, NoHide ; Update the variables with control contents
    SearchTerm := SearchEdit
    if (StrLen(SearchTerm) >= 3) ; Only search if at least three characters have been entered
    {
        SearchFiles(directoryPath, SearchTerm) ; Start the search
    }
return

; Open the selected file when double-clicked or when Enter is pressed
MyListViewDoubleClick:
    if (A_GuiEvent = "DoubleClick") {
        LV_GetText(selectedFileName, "C", A_EventInfo) ; Get the selected filename from the ListView
        selectedFilePath := directoryPath . "\" . selectedFileName ; Construct the full file path
        MsgBox, The selected file is %selectedFilePath%. ; Display a message box with the selected file path
    }
return

Enter::
    LV_GetText(selectedFileName, LV_GetNext()) ; Get the selected filename from the ListView
    selectedFilePath := directoryPath . "\" . selectedFileName ; Construct the full file path
    MsgBox, The selected file is %selectedFilePath%. ; Display a message box with the selected file path
    Run, %selectedFilePath%
return

GuiClose:
ExitApp