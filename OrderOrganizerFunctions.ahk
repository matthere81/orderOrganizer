#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
#include <Vis2>
#include <UIA_Interface>
#include <UIA_Browser>

forwardSoftwareLicense()
{
    ; Get running instance of Outlook
    outlookApp := ComObjActive("Outlook.Application")                           
    
    ; Get MAPI namespace 
    namespace := outlookApp.GetNamespace("MAPI")                                
    
    ; 6 corresponds to the Inbox
    ; folder := namespace.GetDefaultFolder(6).Folders("Test") ; Test folder
    folder := namespace.GetDefaultFolder(6).Folders("SW Licenses")
    
    ; Get all unread emails
    unreadEmails := folder.Items.Restrict("[Unread] = true")                    
    
    for email in unreadEmails
    {
        ; Get the body of the email
        body := email.Body                                                      
        
        ; Get the subject of the email
        subject := email.Subject                                                
        
        ; Find the alphanumeric string after "SO/PO:"
        RegExMatch(body, "SO/PO:\s*([a-zA-Z0-9-]+)", match)
        ; MsgBox % match1          
        salesInfo := getSalesOrderInfo(match1)

        ; If there was an error getting the data, continue to the next email
        if (salesInfo = "error")
        {
            ; Send ^w ; Close the tab
            MsgBox, % "There was an error getting the data. `n`nPO or SO# is " . match 
            . "`n`nThis may mean there are multiple orders with the same number. Please check manually."
            Send ^w ; Close the tab
            continue
        }
        
        salesOrderOrPurchaseOrder := salesInfo[1]
        salesEmail := salesInfo[2]
        firstName := salesInfo[3]
        shippingCustomer := salesInfo[4]
        salesOrderNumber := salesInfo[5]

        ; ; Check if forwardEmail function is successful
        ; if (!forwardEmail(email, salesOrderOrPurchaseOrder, salesEmail, firstName, shippingCustomer, salesOrderNumber))
        ; {
        ;     MsgBox, % "Error forwarding email. Skipping..."
        ;     continue
        ; }

        forwardEmail(email, salesOrderOrPurchaseOrder, salesEmail, firstName, shippingCustomer, salesOrderNumber)
        
        ; Display a message box with Yes and No buttons
        Sleep 1500
        MsgBox, 4,, Send Email?
        IfMsgBox Yes
        {
            email.UnRead := false ; Mark the email as read
            ; email.Send() ; Send the email
            Send !s ; Send the email
        }
        else
        {
            ; Do nothing
        }
    }
}

getSalesOrderInfo(mySalesOrder) {
    browserExe := "chrome.exe"
    orderInfoAddress := "https://customer-insights.thermofisher.com/cockpit/order"
    cockpitAddress := "customer-insights.thermofisher.com/cockpit/order"
    Run, %browserExe% --force-renderer-accessibility %orderInfoAddress%
    cUIA := new UIA_Browser("ahk_exe " browserExe)
    Sleep 2000
    cUIA.WaitPageLoad()
    cUIA.FindFirstByType("Edit").Click()
    Sleep 250
    Clipboard := mySalesOrder
    Send, ^v
    Sleep 250
    Send {enter}
    cUIA.WaitPageLoad()
    url := cUIA.FindFirstByName("Address and search bar")
    url := url.Value 

    ; Check if the title matches the expected title
    if (url != cockpitAddress)
    {
        ; MsgBox, There was an error pulling the info - please check manually.
        ; MsgBox, URL does not match expected address
        ; Send ^w ; Close the tab
        return "error" ; Return from the function to move on to the next email
    }
    
    ; Find the sales person's name
    salesPersonName := cUIA.FindByPath("1,23")
    salesPersonName := salesPersonName.Name
    ; MsgBox, % "Sales Person Name: " . salesPersonName

    ; Convert the sales person's name to lowercase
    StringLower, salesPersonName, salesPersonName

    ; Find the first and last name using regex
    RegExMatch(salesPersonName, "(\w+)\s+(\w\. )?(\w+)", match)
    StringLower, firstName, match1, T
    ; MsgBox, % "First Name: " . firstName

    ; Create the email address
    salesEmail := match1 "." match3 "@thermofisher.com"

    ; Find the customer name
    shippingCustomer := cUIA.FindByPath("1,27")
    shippingCustomer := shippingCustomer.Name
    shippingCustomer := formatText(shippingCustomer)

    ; Find the SO#
    salesOrderNumber := cUIA.FindByPath("1,25")
    salesOrderNumber := salesOrderNumber.Name
    salesOrderNumber := RegExReplace(salesOrderNumber, "^000", "")


    ; MsgBox % salesEmail 
    ; MsgBox % "Shipping Customer: " . shippingCustomer
    ; MsgBox % "Sales Order Number: " . salesOrderNumber

    Sleep 500

    salesInfo := [mySalesOrder, salesEmail, firstName, shippingCustomer, salesOrderNumber]
    Send ^w ; Close the tab
    return salesInfo
}

forwardEmail(email, salesOrder, salesEmail, firstName, shippingCustomer, salesOrderNumber)
{
    ; Forward the email
    forward := email.Forward()

    forward.Display() ; Make the email visible

    ; Set the email properties
    forward.To := salesEmail 
    forward.Subject := "Software License for " shippingCustomer " SO# " salesOrderNumber
    sleep 100
    myBodyOne := "Hi " . firstName . ","
    sleep 100
    myBodyTwo := "Please find the license attached for " shippingCustomer " - SO{#} " . salesOrderNumber . "."
    sleep 100
    myBodyThree := "Thank you"
    sleep 100
    ; forward.Body := myBody . forward.Body

    ; Set focus on the body element
    pasteEmailBody(myBodyOne, myBodyTwo, myBodyThree)
}

pasteEmailBody(myBodyOne, myBodyTwo, myBodyThree)
{
    SetTitleMatchMode, 2
    ControlFocus, _WwG1, Message (HTML), Message
    Clipboard := myBodyOne . "`n`n" . myBodyTwo . "`n`n" . myBodyThree
    Send, % Clipboard
    ; Clipboard := myBodyOne
    ; Sleep 250
    ; Clipwait 1
    ; Send %Clipboard%
    ; Send {Enter 2}
    ; Clipboard := myBodyTwo
    ; Clipwait 1
    ; Send %Clipboard%
    ; Send {Enter 2}
    ; Clipboard := myBodyThree 
    ; Clipwait 1
    ; Send %Clipboard%
}

formatText(shippingCustomer)
{
    ; Split the text into words
    shippingCustomer := StrSplit(shippingCustomer, " ")
    ; Process each word
    for index, word in shippingCustomer
    {
        ; Check if the word has more than three characters
        if (StrLen(word) > 3)
        {
            ; Convert the word to title case
            StringLower, word, word, T
        }
        else
        {
            ; Convert the word to upper case
            StringUpper, word, word
        }

        ; Update the word in the array
        shippingCustomer[index] := word
        ; MsgBox, % shippingCustomer[index]
    }

    ; Join the words back together
    formattedText := ""
    for index, word in shippingCustomer
    {
        ; Add a space before each word except the first one
        if (index > 1)
            formattedText .= " "
        formattedText .= word
    }
    return formattedText
}

forwardSoftwareLicense()
; mySalesOrder := "PUR00688639"
; getSalesOrderInfo(mySalesOrder)