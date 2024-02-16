#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
#include <Vis2>
#include <UIA_Interface>
#include <UIA_Browser>

forwardSoftwareLicense() {
    outlookApp := ComObjActive("Outlook.Application") ; Get running instance of Outlook
    namespace := outlookApp.GetNamespace("MAPI") ; Get MAPI namespace
    
    ; folder := namespace.GetDefaultFolder(6).Folders("SW Licenses") ; 6 corresponds to the Inbox
    folder := namespace.GetDefaultFolder(6).Folders("Test")
    
    unreadEmails := folder.Items.Restrict("[Unread] = true") ; Get all unread emails
    
    for email in unreadEmails {
        body := email.Body ; Get the body of the email
        subject := email.Subject ; Get the subject of the email
        RegExMatch(body, "SO/PO:\s*([a-zA-Z0-9]+)", match) ; Find the alphanumeric string after "SO/PO:"
        ; salesPersonName := getSalesOrderInfo(match1)
        ; forwardEmail(email, salesOrder, salesEmail, firstName)
        salesInfo := getSalesOrderInfo(match1)

        ; MsgBox, % match1 . "`n" . salesInfo[1] . "`n" . salesInfo[2] . "`n" . salesInfo[3]
        salesOrderOrPurchaseOrder := salesInfo[1]
        salesEmail := salesInfo[2]
        firstName := salesInfo[3]
        forwardEmail(email, salesOrderOrPurchaseOrder, salesEmail, firstName)
    }
}

getSalesOrderInfo(mySalesOrder) {
    browserExe := "chrome.exe"
    Run, %browserExe% --force-renderer-accessibility "https://customer-insights.thermofisher.com/cockpit/order"
    cUIA := new UIA_Browser("ahk_exe " browserExe)
    Sleep 2000
    cUIA.WaitElementExistByPath("1.4")
        ; if link := https://customer-insights.thermofisher.com/cockpit/exception/multipleData
        ; {
            
        ; }
    poSearch := cUIA.WaitElementExistByNameAndType("Search by SO/PO", "Edit")
    poSearch.Click()
    Clipboard := mySalesOrder
    Send, ^v{enter}
    cUIA.WaitPageLoad()
    salesPersonName := cUIA.FindByPath("1,23")
    salesPersonName := salesPersonName.Name

    ; Convert the sales person's name to lowercase
    ; StringLower, OutputVar, InputVar , T
    StringLower, salesPersonName, salesPersonName

    ; Find the first and last name using regex
    RegExMatch(salesPersonName, "(\w+)\s+(\w+)", match)
    StringLower, firstName, match1, T

    ; Create the email address
    salesEmail := match1 "." match2 "@thermofisher.com"
    
    salesInfo := [mySalesOrder, salesEmail, firstName]
    return salesInfo
}

forwardEmail(email, salesOrder, salesEmail, firstName)
{
    ; Forward the email
    forward := email.Forward()

    forward.Display() ; Make the email visible

    ; Set the email properties
    forward.To := salesEmail 
    forward.Subject := "Software License for SO# " salesOrder

    myBodyOne := "Hi " . firstName . ","
    myBodyTwo := "Please find the license attached for SO{#} " . salesOrder . "."
    myBodyThree := "Thank you"
    ; forward.Body := myBody . forward.Body

    ; Set focus on the body element
    pasteEmailBody(myBodyOne, myBodyTwo, myBodyThree)
}

pasteEmailBody(myBodyOne, myBodyTwo, myBodyThree)
{
    SetTitleMatchMode, 2
    ControlFocus, _WwG1, Message (HTML), Message
    Clipboard := myBodyOne
    Sleep 250
    Clipwait 1
    Send %Clipboard%
    Send {Enter 2}
    Clipboard := myBodyTwo
    Clipwait 1
    Send %Clipboard%
    Send {Enter 2}
    Clipboard := myBodyThree 
    Clipwait 1
    Send %Clipboard%
}

forwardSoftwareLicense()
