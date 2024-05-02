#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
#include <UIA_Interface>
#include <UIA_Browser>

getQuoteInfo(ByRef quoteID, ByRef contactName, ByRef contactEmail, ByRef contactPhone, ByRef customerName, ByRef quoteOwner, ByRef creatorManager, ByRef totalNetAmount, ByRef totalFreight, ByRef surcharge, ByRef totalTax, ByRef quoteTotal, ByRef soldToID, ByRef paymentTerms, ByRef opportunity)
{
    SetTitleMatchMode 2
    
    if (quoteId = "")
    {
        InputBox, userInput, CPQ#, Please enter CPQ#.
        if ErrorLevel  ; If the user clicked Cancel or didn't enter anything
        {
            MsgBox, CANCEL was pressed or no input was given.
            Return
        }
        userInput := Trim(userInput)
    }
    else
    {
        userInput := quoteId
    }
    
    
    runSalesforceSearch(cUIA, inSearch)
    inSearch.SetValue(userInput)

    ; -------- Click on edit and open opportunity in new window -------- ;
    cUIA.WaitElementExistByName("Suggestions")
    Sleep 750
    preview := cUIA.FindFirstByName("Suggestions")
    preview := preview.FindByPath("1")
    preview.SetFocus()
    Sleep 750
    cUIA.WaitElementExistByName("Record Preview")
    
    ; -------- Get the view link and RegEx it to edit -------- ;
    editLink := cUIA.FindFirstByName("Record Preview")
    editLink := editLink.FindByPath("1")
    editLink := editLink.FindByPath("2")
    editLink := editLink.Value

    ; --------- Set the edit link
    editLink := RegExReplace(editLink, "view$", "edit")    

    ; -------- Navigate to Edit URL -------- ;
    cUIA.SetURL(editLink, True)

    ; -------- Get Info From Salesforce -------- ;
    cUIA.WaitElementExistByName("Quote Information")
    Sleep 500
    quoteID := cUIA.FindFirstByNameAndType("Quote ID", "edit")
    contactName := cUIA.FindFirstByNameAndType("Contact Name", "edit")
    ; contactAddress := cUIA.FindFirstByNameAndType("Contact Address", "edit")
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
    ; contactAddress := contactAddress.Value
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
    opportunityID := opportunity.Value

    quoteID := StrReplace(quoteID, "CPQ-")    
    totalNetAmount := StrReplace(totalNetAmount, "$")
    totalFreight := StrReplace(totalFreight, "$")
    surcharge := StrReplace(surcharge, "$")
    totalTax := StrReplace(totalTax, "$")
    quoteTotal := StrReplace(quoteTotal, "$")

    ; -------- Customer Details Tab -------- ;
    cUIA.FindFirstByName("Customer Details").Click()
    cUIA.WaitElementExistByName("Sold To ID")
    soldToID := cUIA.FindFirstByNameAndType("Sold To ID", "edit")
    soldToID := soldToID.Value
    soldToID := RegExReplace(soldToID, "^0+", "")

    ; -------- Pricing Details -------- ;
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

    if InStr(paymentTerms, "intercompany")
        paymentTerms := "CINC"


}

getWinForm(ByRef opportunity, ByRef winFormLink, ByRef endUser, ByRef endUserPhoneNumber, ByRef endUserEmail, ByRef endUse)
{
    ;-------- GET WIN FORM --------
    MsgBox 36, Get WIN Form?, Get WIN Form?
    IfMsgBox Yes
    {
        runSalesforceSearch(cUIA, inSearch)
        
        ; -------- Set opportunity in search field -------- ;
        inSearch.SetValue(opportunity)
        cUIA.WaitElementExistByName("Suggestions")
        Sleep 750
        winForm := cUIA.FindFirstByName("Suggestions")
        winForm := winForm.FindByPath("1")
        winForm.SetFocus()
        Sleep 750
        cUIA.WaitElementExistByName("Record Preview")
        winForm := cUIA.FindFirstByName("Record Preview")
        winForm.SetFocus()
        winForm := winForm.FindByPath("1")
        winForm := winForm.FindByPath("2")
        winFormLink := winForm.Value
        ; MsgBox % winForm.Value
       
        pattern := "r/(.*)/view"
        RegExMatch(winFormLink, pattern, match)
        match1 := Trim(match1)
        winFormLink := cUIA.SetURL("https://tfs-3--c.vf.force.com/apex/porder_createWIN_p?oppId=" . match1 . "&", True)
        cUIA.WaitPageLoad(winFormLink)

        endUser := cUIA.FindFirstByName("End User Contact Name")
        endUser := endUser.FindByPath("+1")

        endUserPhoneNumber := cUIA.FindFirstByName("End User Phone#")
        endUserPhoneNumber := endUserPhoneNumber.FindByPath("+1")

        endUserEmail := cUIA.FindFirstByName("End User Email")
        endUserEmail := endUserEmail.FindByPath("+1")

        endUse := cUIA.FindFirstBy("AutomationID=thepage\:theform\:j_id109")

        endUser := endUser.Name
        endUserPhoneNumber := endUserPhoneNumber.Name
        endUserEmail := endUserEmail.Name
        endUse := endUse.Value


        ; MsgBox % endUser . "`n" . endUserPhoneNumber . "`n" . email . "`n" . endUse
    }
    Else
    {
        Return
    }
}

getQuoteInfo(quoteID, contactName, contactEmail, contactPhone, customerName, quoteOwner, creatorManager, totalNetAmount, totalFreight, surcharge, totalTax, quoteTotal, soldToID, paymentTerms, opportunity)


getWinForm(opportunity, winFormLink, endUser, endUserPhoneNumber, endUserEmail, endUse)


runSalesforceSearch(ByRef cUIA, ByRef inSearch)
{
    browserExe := "chrome.exe"
    Run, %browserExe% --force-renderer-accessibility "https://tfs-3.lightning.force.com/lightning/page/home"
    ; Run, %browserExe% --force-renderer-accessibility "https://tfs-3.lightning.force.com/lightning/r/cafsl__Oracle_Quote__c/a5B4z000000ku6hEAA/edit"
    
    cUIA := new UIA_Browser("ahk_exe " browserExe)
    Sleep 2000
    cUIA.WaitElementExistByPath("1.4")
    search := cUIA.FindByPath("1.4")
    Sleep 750

    ; -------- Set quote# in search field -------- ;
    search.Click()
    inSearch := cUIA.WaitElementExistByNameAndType("Search...", "Edit")
    inSearch.Click()
}