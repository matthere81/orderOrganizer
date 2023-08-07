#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
#include <UIA_Interface>
#include <UIA_Browser>

getQuoteInfo(ByRef quoteID, ByRef contactName, ByRef contactAddress, ByRef contactEmail, ByRef contactPhone, ByRef customerName, ByRef quoteOwner, ByRef creatorManager, ByRef totalNetAmount, ByRef totalFreight, ByRef surcharge, ByRef totalTax, ByRef quoteTotal, ByRef soldToID, ByRef paymentTerms, ByRef opportunity)
{
    SetTitleMatchMode 2

    InputBox, userInput, CPQ#, Please enter CPQ#.
    if ErrorLevel  ; If the user clicked Cancel or didn't enter anything
    {
        MsgBox, CANCEL was pressed or no input was given.
        Return
    }
    
    browserExe := "chrome.exe"
    ; Run, %browserExe% --force-renderer-accessibility "https://tfs-3.lightning.force.com/lightning/page/home"
    Run, %browserExe% --force-renderer-accessibility "https://tfs-3.lightning.force.com/lightning/r/cafsl__Oracle_Quote__c/a5B4z000000ku6hEAA/view"
    
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

    ; -------- Get Info From Salesforce -------- ;
    cUIA.WaitElementExistByName("Quote Information")
    Sleep 500
    quoteID := cUIA.FindFirstByNameAndType("Quote ID", "edit")
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
    opportunity := cUIA.FindFirstByNameAndType("Opportunity", "edit")
    
    quoteID := quoteID.Value
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
    opportunity := opportunity.Value

    quoteID := StrReplace(quoteID, "CPQ-")    
    totalNetAmount := StrReplace(totalNetAmount, "$")
    totalFreight := StrReplace(totalFreight, "$")
    surcharge := StrReplace(surcharge, "$")
    totalTax := StrReplace(totalTax, "$")
    quoteTotal := StrReplace(quoteTotal, "$")

    ; Customer Details Tab
    cUIA.FindFirstByName("Customer Details").Click()
    cUIA.WaitElementExistByName("Sold To ID")
    soldToID := cUIA.FindFirstByNameAndType("Sold To ID", "edit")
    soldToID := soldToID.Value
    soldToID := RegExReplace(soldToID, "^0+", "")

    ; Pricing Details
    cUIA.WaitElementExistByName("Pricing Details")
    cUIA.FindFirstByName("Pricing Details").Click()
    cUIA.WaitElementExistByName("Payment Terms")
    Sleep 500
    Send {tab}
    Sleep 250
    Send ^c
    Sleep 250
    paymentTerms := Clipboard
    paymentTerms := RegExReplace(paymentTerms, "i)days.*", "Days")
}

getQuoteInfo(quoteID, contactName, contactAddress, contactEmail, contactPhone, customerName, quoteOwner, creatorManager, totalNetAmount, totalFreight, surcharge, totalTax, quoteTotal, soldToID, paymentTerms, opportunity)