#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
#include <UIA_Interface>
#include <UIA_Browser>

SetTitleMatchMode 2

userInput := 00592335
; InputBox, userInput, CPQ#, Please enter CPQ#.
; if ErrorLevel ; If the user clicked Cancel or didn't enter anything
;     MsgBox, CANCEL was pressed or no input was given.

browserExe := "chrome.exe"
; Run, %browserExe% --force-renderer-accessibility "https://tfs-3.lightning.force.com/lightning/page/home"
Run, %browserExe% --force-renderer-accessibility "https://tfs-3.lightning.force.com/lightning/r/a5B4z000000l5pFEAQ/edit"
cUIA := new UIA_Browser("ahk_exe " browserExe)
Sleep 2000
; cUIA.WaitPageLoad()
; Sleep 1000
; cUIA.WaitPageLoad()
; search := cUIA.FindByPath("1.4")
; Sleep 500

; ; -------- Set quote# in search field -------- ;
; search.Click()
; inSearch := cUIA.WaitElementExistByNameAndType("Search...", "Edit")
; inSearch.Click()
; inSearch.SetValue(userInput)

; ; -------- Click on edit and open opportunity in new window -------- ;
; cUIA.WaitElementExistByName("Suggestions")
; Sleep 500
; preview := cUIA.FindByPath("1.19.7.1")
; preview.SetFocus()
; Sleep 500
; cUIA.WaitElementExistByName("Record Preview")
; editLink := cUIA.FindByPath("1.19.9.1.2")

; ; -------- Get the view link and RegEx it to edit -------- ;
; editLink := editLink.Value
; editLink := RegExReplace(editLink, "view$", "edit")

; ; -------- Navigate to Edit URL -------- ;
; cUIA.SetURL(editLink, True)


cUIA.WaitElementExistByName("Quote Information")
Sleep 500
quoteId := cUIA.FindFirstByNameAndType("Quote ID", "edit")
contactName := cUIA.FindFirstByNameAndType("Contact Name", "edit")
contactAddress := cUIA.FindFirstByNameAndType("Contact Address", "edit")
contactEmail := cUIA.FindFirstByNameAndType("Contact Email", "edit")
contactPhone := cUIA.FindFirstByNameAndType("Contact Phone", "edit")
customerName := cUIA.FindFirstByNameAndType("Account Name", "edit")
quoteOwner := cUIA.FindFirstByNameAndType("Owner", "edit")
creatorManager := cUIA.FindFirstByNameAndType("Creator Manager", "edit")
totalNetAmount := cUIA.FindFirstByNameAndType("Total Net Amount", "edit")
totalFreight := cUIA.FindFirstByNameAndType("Total Freight", "edit")
surcharge := cUIA.FindFirstByNameAndType("Surcharge", "edit")
totalTax := cUIA.FindFirstByNameAndType("Total Tax / VAT / GST", "edit")
quoteTotal := cUIA.FindFirstByNameAndType("Quote Total", "edit")

MsgBox % quoteId.Value . "`n" . contactName.Value . "`n" . contactAddress.Value . "`n" . contactEmail.Value
    . "`n" . contactPhone.Value . "`n" . customerName.Value . "`n" . quoteOwner.Value . "`n" . creatorManager.Value
    . "`n" . totalFreight.Value . "`n" . surcharge.Value . "`n" . totalTax.Value . "`n" . quoteTotal.Value

Return

Send +{F10}
Sleep 100
Send {Down}{Enter}
WinActivate Home | Salesforce
Sleep 1000
; cUIA.FindByPath("1.19.9.2.1.1").Highlight()
currentProperty := cUIA.FindByPath("1.19.9.2.1.1").GetCurrentDocumentElement()
MsgBox % currentProperty

; Get edit link
; Download to URL

Return
cUIA.FindByPath("1.19.9.1.2").SetFocus()
Sleep 1000
; cUIA.FindByPath("1.19.9.2.1.1").Click()
cUIA.FindByPath("1.19.9.2.1.1").Highlight()
cUIA.WaitElementExistByName("Quote Information")

Sleep 1000

; -------- DOES THE DUMP -------- ;
Clipboard := cUIA.DumpAll()

Return

isSalesforceOpen()
{
    SetTitleMatchMode 2
    ; Set target tab name
    vTitle := "Salesforce"
    DetectHiddenWindows, Off

    ; Get list of chrome tabs
    WinGet vList, List, % "ahk_exe chrome.exe"
    ; cycle through all chrome windows
    Loop % vList
    {
        WinActivate % "ahk_id" vList%A_Index%
        WinWaitActive % "ahk_id" vList%A_Index%
        Sleep 25
        hwnd := vList%A_Index%

        ; Cycle through all tabs
        Loop 100
        {
            WinGetTitle, title_chromeTab, ahk_id %hwnd%
            
            ; quit this window if circled all the way back to first seen tab
            if (title_chromeTab == firstTab)
            {
                ; msgbox, % "looped through " A_Index "Tabs."
                break
            }

            ; Remember first tab so we know where to stop
            if (A_Index == 1)
            {
                firstTab := title_chromeTab
            }

            ; If this is the window we are looking for
            if (InStr(title_chromeTab, vTitle))
            {
                msgbox, % "Found " vTitle " for you!"
                WinActivate title_chromeTab
                break 2
            }
            ; proceed to next tab
            Send ^{Tab}
            sleep, 100
            ; msgbox, % title_chromeTab
        }
    }
    Return
}
Return