#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
#include <UIA_Interface>
#include <UIA_Browser>

getQuoteInfo(ByRef quoteId, ByRef contactName, ByRef contactAddress, ByRef contactEmail, ByRef contactPhone, ByRef customerName, ByRef quoteOwner, ByRef creatorManager, ByRef totalNetAmount, ByRef totalTax, ByRef totalFreight, ByRef surcharge, ByRef quoteTotal)
{
    SetTitleMatchMode 2

    InputBox, userInput, CPQ#, Please enter CPQ#.
    if ErrorLevel  ; If the user clicked Cancel or didn't enter anything
    {
        MsgBox, CANCEL was pressed or no input was given.
        Return
    }

    browserExe := "chrome.exe"
    Run, %browserExe% --force-renderer-accessibility "https://tfs-3.lightning.force.com/lightning/page/home"
    cUIA := new UIA_Browser("ahk_exe " browserExe)
    Sleep 2000
    cUIA.WaitElementExistByPath("1.4")
    search := cUIA.FindByPath("1.4")
    Sleep 500

    ; -------- Set quote# in search field -------- ;
    search.Click()
    inSearch := cUIA.WaitElementExistByNameAndType("Search...", "Edit")
    inSearch.Click()
    inSearch.SetValue(userInput)

    ; -------- Click on edit and open opportunity in new window -------- ;
    cUIA.WaitElementExistByName("Suggestions")
    Sleep 500
    preview := cUIA.FindByPath("1.19.7.1")
    preview.SetFocus()
    Sleep 500
    cUIA.WaitElementExistByName("Record Preview")
    editLink := cUIA.FindByPath("1.19.9.1.2")

    ; -------- Get the view link and RegEx it to edit -------- ;
    editLink := editLink.Value
    editLink := RegExReplace(editLink, "view$", "edit")

    ; -------- Navigate to Edit URL -------- ;
    cUIA.SetURL(editLink, True)


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

    quoteId := quoteId.Value
    quoteId := StrReplace(quoteId, "CPQ-")
    contactName :=contactName.Value
    contactAddress := contactAddress.Value
    contactEmail := contactEmail.Value
    contactPhone := contactPhone.Value
    customerName := customerName.Value
    quoteOwner := quoteOwner.Value
    creatorManager := creatorManager.Value
    totalNetAmount := totalNetAmount.Value
    totalFreight := totalFreight.Value
    surcharge := surcharge.Value
    totalTax := totalTax.Value
    quoteTotal := quoteTotal.Value
    totalNetAmount := StrReplace(totalNetAmount, "$")
    totalFreight := StrReplace(totalFreight, "$")
    surcharge := StrReplace(surcharge, "$")
    totalTax := StrReplace(totalTax, "$")
    quoteTotal := StrReplace(quoteTotal, "$")
    MsgBox % totalNetAmount . " is totalNetAmount at end of QuoteInfo.ahk"
    ; Return quoteId, ;contactName, contactAddress, 
    ; MsgBox % "Quote# is " . quoteId . "`nContact Name is " . contactName . "`nAddress is " . contactAddress . "`nEmail is " . contactEmail
    ;     . "`nPhone is " . contactPhone . "`nCustomer is " . customerName . "`nSalesperson is " . quoteOwner . "`nManager is " . creatorManager
    ;     . "`nFreight is " . totalFreight . "`nsurcharge is " . surcharge . "`ntax is " . totalTax . "`nQuote Total is " . quoteTotal

}

