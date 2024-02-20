#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%


SearchFilesFunction(SearchBox)
{
    Gui, Submit, NoHide
    searchTerm := SearchBox

    ; If the search box is empty, hide the ListView and return
    if (searchTerm = "")
    {
        GuiControl, Hide, FileList
        Gui, Show, NoActivate  ; Re-draw the GUI
        return
    }
    else
    {
        GuiControl, Show, FileList
        Gui, Show, NoActivate  ; Re-draw the GUI
    }

    ; Clear the ListView before adding new items
    LV_Delete()

    Loop, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB\*.*
    {
        If InStr(A_LoopFileName, searchTerm)
        {
            ; Add the file name to the ListView
            LV_Add("", A_LoopFileName)
        }
    }
}

; from copilot:
; Define functions for repetitive tasks
; AddText(text, variable) {
;     Gui Add, Text,, %text%
;     Gui Add, Edit, v%variable%, % %variable% %
; }

; AddGroupBox() {
;     Gui Add, GroupBox, x+12.5 y100 w1 h400 ; vertical line
; }

; ; Use an associative array for sales directors and their codes
; salesDirectors := {"Denise Schwartz": "201020", "Joann Purkerson": "202375", "Maroun El Khoury": "96715", "Jimmy Yuk": "1261", "Sylveer Bergs": "98866", "N/A": "N/A"}

; ; Externalize hard-coded values
; myinipath = %A_MyDocuments%\Order Docs\Info DB
; I_Icon = %A_Desktop%\Auto Hot Key Scripts\list_check_checklist_checkmark_icon_181579.ico

; ; Use the functions to add elements to the GUI
; AddText("CPQ:", "cpq")
; AddText("PO:", "po")
; AddGroupBox()
; AddText("System:", "system")
; AddText("Salesperson:", "salesPerson")