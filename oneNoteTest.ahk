#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

; Create a new OneNote application object
oneNoteApp := ComObjCreate("OneNote.Application")

; Get the XML content of the specified note
noteLink := "onenote:https://thermofisher-my.sharepoint.com/personal/matthew_terbeek_thermofisher_com/_layouts/OneNote.aspx?id=%2Fpersonal%2Fmatthew_terbeek_thermofisher_com%2FDocuments%2FMatthew%20%40%20Thermo%20Fisher%20Scientific&wd=target%28Order%20Workflow.one%7C188B4DAA-9ACE-49D8-81B8-E9A0274C5165%2FE-mail%7CFF412216-1FF8-4D3D-89A6-B0FD72271307%2F%29"
noteXml := oneNoteApp.GetPageContent(noteLink)

; Extract the text content from the XML
xmlDoc := ComObjCreate("MSXML2.DOMDocument")
xmlDoc.loadXML(noteXml)

; Extract the text content from the XML using XPath
textNode := xmlDoc.selectSingleNode("//one:Page/one:Outline/one:OEChildren/one:OE/one:T")
noteContent := textNode.text

; Display the note content
MsgBox % noteContent

; Release the OneNote application object
oneNoteApp := ""

Return