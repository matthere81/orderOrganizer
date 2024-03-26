Search:
Gui Submit, NoHide ; Update the variables with the current values of the controls
SearchTerm := Trim(SearchTerm) ; Remove trailing whitespace from SearchTerm

Loop, %myinipath%\*.* ; Loop through all files in the selected folder
{
    IfInString, A_LoopFileName, %SearchTerm% ; If the current file name contains the search term
        MsgBox, %A_LoopFileName% ; Display a message box with the file name
}

matchingFiles := [] ; Create an empty array to store matching files

Loop, %myinipath%\\*.* ; Loop through all files in the selected folder
{
    IfInString, A_LoopFileName, %SearchTerm% ; If the current file name contains the search term
    {
        matchingFiles.Push(A_LoopFileName) ; Add the file name to the list of matching files
    }
}

if (matchingFiles.Length() = 1) ; If there's only one matching file
    {
        SelectedFile := matchingFiles[1] ; Set the 'file' variable to the name of the file
        Gosub readtheini
    }
    else if (matchingFiles.Length() > 1) ; If there's more than one matching file
    {
        fileList := "" ; Create an empty string to store the file names
        Loop, % matchingFiles.Length() ; Loop over the 'matchingFiles' array
        {
            fileList := fileList . matchingFiles[A_Index] . "`n" ; Append the file name to the string
        }
        MsgBox, % "Please choose a file from the following list:`n" . fileList ; Display the file names in a MsgBox
    }

return