

;----------- START CHECKLISTS ---------------

Gui Add,Tab3,, Salesforce Checklist|SAP Checklist - Main Page|SAP Checklist - Inside The Order|SAP - Finalizing The Order
Gui Tab, 1

;|------------------------------------------|
;|                                          |
;| Define the section 1 and its checklists  |
;|                                          |
;|------------------------------------------|

section1 := ["Pre Salesforce", "Salesforce", "Pre SAP"]
checklists1 := [["Check TENA name on PO", "Quote number matches on PO && quote", "Payment terms match on PO && quote", "Prices match on PO && quote"
                , "Bill to && Ship to address on PO"], ["Save PDF of Quote (Document Output Tab)", "Arrange quote lines (if needed) (Quote Details Tab)"
                , "Sold-To ID (Customer Details Tab)", "Check order type (Order Tab)", "Add PO# / Add PO Value / Upload PO (Order Tab)"]
                , ["Order Notice Sent", "Entered In SOT", "T&&Cs?"]]

; Loop over each section
Loop, % section1.MaxIndex()
{
    index := A_Index
    ; Call the subroutine with the current section index, section and checklist
    AddSection(section1[A_Index], checklists1[A_Index])
}


;|------------------------------------------|
;|                                          |
;| Define the section 2 and its checklists  |
;|                                          |
;|------------------------------------------|

Gui Tab, 2

section2 :=["SAP Attachments", "Main Sales Tab"]
checklists2 := [["Purchase Order", "CPQ Quote", "Order Notice To Sales", "WIN Form", "Merge Report"], ["CRD Date", "First Date Lines Updated", "Incoterms"]]

; Loop over each section
Loop, % section2.MaxIndex()
{
    index := A_Index
    ; Call the subroutine with the current section index, section and checklist
    AddSection(section2[A_Index], checklists2[A_Index], true)
}

;|------------------------------------------|
;|                                          |
;| Define the section 3 and its checklists  |
;|                                          |
;|------------------------------------------|

Gui Tab, 3

section3 :=["Inside the Order", "Partners Tab", "Texts Tab"]
checklists3 := [["110 Volts (Sales Tab)", "Shipper (Shipping Tab)", "Verify Incoterms (Billing Tab)"
                ,"Check for Partner", "Check for Partner Function"], ["Manager Code (Sales Manager)"
                ,"Director Code (Sales Employee 9)", "Verify Billing Address", "Verify Shipping Address"
                ,"End User Added", "Contact Person Added"], ["Contact Person or End User Info Added"
                ,"Serial/Dongle Number Added To Form Header?", "End User info added in the partners tab?"]]

                
; Loop over each section
Loop, % section3.MaxIndex()
{
    index := A_Index
    ; Call the subroutine with the current section index, section and checklist
    AddSection(section3[A_Index], checklists3[A_Index], true)
}

;|------------------------------------------|
;|                                          |
;| Define the section 4 and its checklists  |
;|                                          |
;|------------------------------------------|

Gui Tab, 4

section4 := ["Finalizing the Order"]
checklists4 := [["Check Net Value", "Add Shipping?", "Higher Level Linking?", "Delivery Groups?"
                ,"Update Delivery Block", "Order Acceptance"]]

; Loop over each section
Loop, % section4.MaxIndex()
{
    index := A_Index
    ; Call the subroutine with the current section index, section and checklist
    AddSection(section4[A_Index], checklists4[A_Index], true)
}

; Create an array of all checklist arrays
allChecklists := [checklists1, checklists2, checklists3, checklists4]

; ---------- END PRE-SAP ----------