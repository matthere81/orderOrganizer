#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
;~ #Include WatchFolder.ahk
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir, C:\Users\%A_UserName%\Order Organizer ; Ensures a consistent starting directory.
#SingleInstance Force

salesPeople := "|Justin Carder|Robin Sutka|Fred Simpson|Rhonda Oesterle|Mitch Lazaro|Tucker Lincoln|Jawad Pashmi|Julie Sawicki|Mike Hughes|Steve Boyanoski"
. "|Bob Riggs|Chuck Costanza|Navette Shirakawa|Stephanie Koczur|Mark Krigbaum|Jon Needels|Bill Balsanek|Brent Boyle|Andrew Clark|Kevin Clodfelter|Gabriel Mendez"
. "|Karl Kastner|Michael Burnett|Jerry Pappas|Nick Duczak|Steven Danielson|Nick Hubbard|Samantha Stikeleather|Drew Smillie|Jeff Weller|Jerry Holycross"
. "|Theresa Borio|Dan Ciminelli|Cynthia Spittler|Gwyn Trojan|Joel Stradtner|Don Rathbauer|Hillary Tennant|Melissa Chandler|Douglas Sears|Rashila Patel|Brian Thompson"
. "|Larry Bellan|Donna Zwirner|Kristen Luttner|Helen Sun|May Chou|Haris Dzaferbegovic|Brian Dowe|Mark Woodworth|Susan Bird|Giovanni Pallante|Alicia Arias"
. "|Dominique Figueroa|Jonathan McNally|Murray Fryman|Yan Chen|Jie Qian|Joe Bernholz|David Kage|David Scott|Todd Stoner|John Bailey|Katianna Pihakari|Jonathan Ferguson"
. "|Aeron Avakian|Luke Marty|Alexander James|Timothy Johnson|Yuriy Dunayevskiy|Susan Gelman|Cari Randles|Shijun Sheng|Sean Bennett|Nelson Huang|Lorraine Foglio|Gerald Koncar"
. "|Lauren Fischer|Brian Luckenbill|Amy Allgower|Brandon Markle|Crystal Flowers|Douglas McDowell|"

salesManagers := "|Anjou Keller|Joe Hewitt|Zee Nadjie|Doug McCormack|Natalie Foels|Tonya Second|Lou Gavino|Christopher Crafts|Joe McFadden|John Butler|Richard Klein|Ray Chen|Randy Porch"

salesDirectors := "|Joann Purkerson|Maroun El Khoury|Jimmy Yuk|N/A"

salesCodes := "|202375|96715|1261|98866|96695|96654|202625|202006|1076|95410|202610|1026|1042|202756|202611|1041|N/A"

; Set Order Organizer Path
if !FileExist("C:\Users\" . A_UserName . "\Order Organizer") {
    FileCreateDir, A_WorkingDir . "\Order Organizer"
    SetWorkingDir, A_WorkingDir . "\Order Organizer"
}

myinipath := "C:\Users\" . A_UserName . "\Order Organizer\Order Database"
if !FileExist(myinipath) {
    FileCreateDir, %myinipath%
}

; Include Icon
FileInstall, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Desktop\Auto Hot Key Scripts\list_check_checklist_checkmark_icon_181579.ico, A_WorkingDir, 1

SetTitleMatchMode, 2

orderInfo(){
    global
;/******** GUI START ********\
Gui destroy
Gui Font
Gui Font, s12 w600 Italic cBlack, Tahoma
Gui Add, Text, x10 y30, ______________________________________________________________
Gui Add, Text, hWndhTxtOrderOrganizer23 x15 y20 w300 +Left, Order Organizer ; - SO# %soNumber%
Gui Font
Gui Color, 79b8d1
Gui Font, S9, Segoe UI Semibold
Gui Add, Button, x650 y20 w70 greadtheini, O&pen
Gui Add, Button, x+25 w70 gSaveToIni, &Save
Gui Add, Button, x+25 w125 grestartScript, &New PO or Reload
; Gui Add, Edit, x+25 y22 w175 h20, Search

;----------- COLUMN 1 ---------------

Gui Add, Text, Section x50 y70, SOT Line#
Gui Add, Edit, yp+20 xp-2.5 vsot, %sot%
Gui Add, Text,, Customer:
Gui Add, Edit, vcustomer, %customer% 
Gui Add, Text,, Customer / DPS Contact:
Gui Add, Edit, vcontact, %contact% 
Gui Add, Text,, Customer / DPS Address:
Gui Add, Edit, vaddress, %address% 
Gui Add, Text,, Sold To Account:
Gui Add, Edit, vsoldTo, %soldTo%
Gui Add, Text,, SO#
Gui Add, Edit, vsoNumber, %soNumber% 

;----------- END COLUMN 1 END ---------------

;----------- COLUMN 2 ---------------
Gui Add, Text, ys x+45 Section, CPQ or Quote#
Gui Add, Edit, yp+20 xp-2.5 vcpq, %cpq%
Gui Add, Text,, System:
Gui Add, Edit, vsystem, %system% 
Gui Add, Text,, CRD (Cust Req Date):
Gui Add, DateTime, w135 vcrd, %crd%
Gui Add, Text,, PO Date:
Gui Add, DateTime, w135 vpoDate, %poDate% 
Gui Add, Text,, SAP Date:
Gui Add, DateTime, w135 vsapDate, %sapDate%
Gui Add, Text,, Software upgrade:
Gui Add, Radio, x+-80 y+10 gsubmitChecklist vsoftwareUpgradeYes, Yes
Gui Add, Radio, x+5 gsubmitChecklist vsoftwareUpgradeNo, N/A

;----------- END COLUMN 2 END ---------------

;----------- COLUMN 3 ---------------

Gui Add, Text, ys x+80 Section, PO#
Gui Add, Edit, yp+20 xp-2.5 vpo, %po%
Gui Add, Text,, PO Value:
Gui Add, Edit, vpoValue, %poValue%
Gui Add, Text,, Tax:
Gui Add, Edit, vtax, %tax%
Gui Add, Text,, Freight Cost:
Gui Add, Edit, vfreightCost, %freightCost% 
Gui Add, Text,, Total:
Gui Add, Edit, vtotalCost, %totalCost% 
Gui Add, Text,, Serial | License | Dongle #:
Gui Add, Edit, vsoftwareUpgradeLicense, %softwareUpgradeLicense% 

;----------- END COLUMN 3 END ---------------

;----------- COLUMN 4 ---------------

Gui, Add, Text, ys x+45 Section, Salesperson:
Gui, Add, DropDownList, yp+22.5 xp-2.5 +Sort vsalesPerson gsubmitSales, % salesPeople
Gui, Add, Text, y+7.5, Sales Manager:
Gui, Add, DDL, Disabled vsalesManager, % salesManagers
Gui Add, Text, y+10, Sales Manager Code:
Gui Add, DropDownList, ReadOnly vmanagerCode, % salesCodes
Gui, Add, Text, y+7.5, Sales Director:
Gui, Add, DropDownList, Disabled vsalesDirector, % salesDirectors
Gui Add, Text, y+10, Sales Director Code:
Gui, Add, DropDownList, ReadOnly vdirectorCode, % salesCodes
Gui Add, Text,, Notes:
Gui Add, Edit, w300 h85 vnotes, %notes%


;----------- END COLUMN 4 END ---------------

;----------- COLUMN 5 ---------------

; Gui Add, Text, ys x+45, SO#
; Gui Add, Edit, yp+20 xp-2.5 vsoNumber, %soNumber% 
; Gui Add, Text, y+10, Software upgrade:
; Gui Add, Radio, gsubmitChecklist vsoftwareUpgradeYes, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vsoftwareUpgradeNo, N/A
; Gui Add, Text, x+-90 y+10, Serial | License | Dongle #:
; Gui Add, Edit, vsoftwareUpgradeLicense, %softwareUpgradeLicense% 

Gui Add, Text, ys x+-135 Section, End User:
Gui Add, Edit, vendUser
Gui Add, Text,, End User / Contact Phone:
Gui Add, Edit, vphone
Gui Add, Text,, End User / Contact Email:
Gui Add, Edit, vemail
Gui Add, Text,, End Use:
Gui Add, Edit, h78 vendUse


;----------- END COLUMN 5 END ---------------
; Gui Font, c0052cf
Gui Add, Text, x10 y450, ________________________________________________________________________________________________________________________________________________________________________________________________
; Gui Font

;----------- END MAIN SECTION ---------------

;----------- START CHECKLISTS ---------------

Gui Add, Tab3, x25 w925 h220, Salesforce Checklist|SAP Checklist (Main Page)|SAP Checklist (Inside The Order)|SAP - Finalizing The Order
Gui Tab, 1


; ;======== KEYBOARD SHORTCUTS ========
; Gui Add, Listview, y+10 w215 h235 R13 grid ReadOnly, Value (Keyboard Shortcut)
; LV_ModifyCol(1,190)
; ;~ LV_ModifyCol(2, 115)
; LV_Add(Col1, "CPQ (zpq)") ;"zpq")
; ;~ Gui Add, Text,, CPQ - zpq
; LV_Add(Col1, "PO# (zpo)") ;,"zpo")
; ;~ Gui Add, Text, y+5, PO# - zpo
; LV_Add(Col1, "SO# (zso)") ;,"zso")
; ;~ Gui Add, Text, y+5, SO# - zso
; LV_Add(Col1, "SOT Line# (zsot)") ;,"zsot")
; ;~ Gui Add, Text, y+5, SOT Line# - zsot
; LV_Add(Col1, "Customer (zcust)") ;,"zcust")
; ;~ Gui Add, Text, y+5, Customer - zcust
; LV_Add(Col1, "PO Value (zval)") ;,"zval")
; ;~ Gui Add, Text, y+5, PO Value - zval
; LV_Add(Col1, "Salesperson (zsal)") ;,"zsal")
; ;~ Gui Add, Text, y+5, Salesperson - zsal
; LV_Add(Col1, "Cust Contact (zcon)") ;,"zcon")
; ;~ Gui Add, Text, x775 y225, Customer Contact -
; ;~ Gui Add, Text, y+5, zcon
; LV_Add(Col1, "Cust Email (zem)")
; ;~ Gui Add, Text, y+5, Customer Email - 
; ;~ Gui Add, Text, y+5, zem
; LV_Add(Col1, "System (zsys)")
; ;~ Gui Add, Text, y+5, System - zsys
; LV_Add(Col1, "End User (zenu)")
; ;~ Gui Add, Text, y+5, End User - zenu
; LV_Add(Col1, "End User Phone (zph)")
; ;~ Gui Add, Text, y+5, Phone - zph
; LV_Add(Col1, "End Use (zuse)")
; ;~ Gui Add, Text, y+5, End Use - zuse

; ;======== END KEYBOARD SHORTCUTS ========

; ------------- CHECKLIST TABS ---------------
; Gui, Tab, 2

; ----------- PRE SALESFORCE -----------------

Gui Add, GroupBox,Section h175 w250, Pre Salesforce
Gui Add, Checkbox, xp+10 yp+30 gsubmitChecklist vnameCheck, Check TENA name on PO
Gui Add, Checkbox, gsubmitChecklist vquoteNumberMatch, Quote number matches on PO && quote
Gui Add, Checkbox, gsubmitChecklist vpaymentTerms, Payment terms match on PO && quote
Gui Add, Checkbox, gsubmitChecklist vpriceMatch, Prices match on PO && quote
Gui Add, Checkbox, gsubmitChecklist vbothAddresses, Bill to && Ship to address on PO

; ----------- END PRE SALESFORCE -----------------

; ----------- SALESFORCE -----------------

Gui Add, GroupBox, Section ys h175 w325, Salesforce
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vpdfQuote, Save PDF of Quote (Document Output Tab)
Gui Add, Checkbox, gsubmitChecklist varrangeLines, Arrange quote lines (Quote Details Tab)
Gui Add, Checkbox, gsubmitChecklist vsoldToIdCheck, Sold-To ID (Customer Details Tab)
Gui Add, Checkbox, gsubmitChecklist vorderTypeCheck, Check order type (Order Tab)
Gui Add, Checkbox, gsubmitChecklist vpoInfoCheck, Add PO# / Add PO Value / Upload PO (Order Tab)
Gui Add, Checkbox, gsubmitChecklist vgenerateDps, Generate && Attach DPS Reports (Attachments Tab)

; ----------- END SALESFORCE -----------------

; ----------- PRE SAP -----------------

Gui Add, GroupBox, Section ys h175 w275, Pre SAP
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vorderNoticeSent, Order Notice Sent
Gui Add, Checkbox, gsubmitChecklist venteredSot, Entered In SOT
Gui Add, Text, tcs y+10, T&&Cs?
Gui Add, Radio, x+5 gsubmitChecklist vtandcYes, Yes
Gui Add, Radio, x+5 gsubmitChecklist vtandcNa, N/A

; ---------- END PRE-SAP ----------

Gui, Tab, 2
; ----------- SAP -----------------

Gui Add, GroupBox,Section x+15 ys+10 h130 w215, SAP Attachments:
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vpoAttached, PO
Gui Add, Checkbox, x+55 gsubmitChecklist vquoteAttached, Quote
Gui Add, Checkbox, xm+35 y+10 gsubmitChecklist vdpsAttached, DPS Reports
Gui Add, Checkbox, x+5 gsubmitChecklist vorderNoticeAttached, Order Notice
Gui Add, Text, xm+35 y+10, WIN Form?
Gui Add, Radio, x+25 gsubmitChecklist vwinYes, Yes
Gui Add, Radio, x+5 gsubmitChecklist vwinNa, N/A
Gui Add, Text, xm+35 y+10, Merge Report?
Gui Add, Radio, x+10 gsubmitChecklist vmergeYes, Yes
Gui Add, Radio, x+5 gsubmitChecklist vmergeNa, N/A

Gui Add, GroupBox, Section ys h100 w290, Main Sales Tab
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vcrdDateAdded, CRD Date (Req delv date)
Gui Add, Checkbox, y+10 gsubmitChecklist vfirstDate, First Date Lines Updated (Edit -> Fast Change)
Gui Add, Checkbox, gsubmitChecklist vincoterms, Incoterms

Gui, Tab, 3

Gui Add, GroupBox, Section h100 w200, Inside the Order
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vvolts, 110 Volts (Sales Tab)
Gui Add, Checkbox, gsubmitChecklist vshipper, Shipper (Shipping Tab)
Gui Add, Checkbox, gsubmitChecklist vverifyIncoterms, Verify Incoterms (Billing Tab)

Gui Add, GroupBox, ys w250 h175 Section, Partners Tab
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vmanagerCodeCheck, Manager Code (Sales Manager)
Gui Add, Checkbox, gsubmitChecklist vdirectorCodeCheck, Director Code (Sales Employee 9)
Gui Add, Checkbox, gsubmitChecklist vbilltoCheck, Verify Billing Address
Gui Add, Checkbox, gsubmitChecklist vshiptoCheck, Verify Shipping Address
Gui Add, Text,, End User Added
Gui Add, Radio, x+25 gsubmitChecklist vendUserCheck, Yes
Gui Add, Radio, x+5 gsubmitChecklist vendUserNA, N/A
Gui Add, Checkbox, xs+10 y+10 gsubmitChecklist vcontactPersonCheck, Contact Person Added

Gui Add, GroupBox, ys Section h175 w275, Texts Tab
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vtextsContactCheck, Contact Person or End User Info Added
Gui Add, Text,, Serial/Dongle Number Added To Form Header?
Gui Add, Radio, xs+25 y+10 gsubmitChecklist vserialYes, Yes
Gui Add, Radio, x+5 gsubmitChecklist vserialNa, N/A
Gui Add, Text, xs+10 y+10, End User info added in the partners tab?
Gui Add, Radio, xs+25 y+10 gsubmitChecklist vendUserCopyBack, Yes
Gui Add, Radio, x+5 gsubmitChecklist vendUserCopyBackNa, N/A

Gui, Tab, 4

Gui Add, GroupBox,Section h175 w235, Final Steps:
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vcheckPrices, Check Net Value
Gui Add, Text,, Add Shipping?
Gui Add, Radio, x+50 gsubmitChecklist vshippingYes, Yes
Gui Add, Radio, x+5 gsubmitChecklist vshippingNa, N/A
Gui Add, Text, xs+10 y+10, Higher Level Linking?
Gui Add, Radio, x+15 gsubmitChecklist vhigherLevelLinkingYes, Yes
Gui Add, Radio, x+5 vhigherLevelLinkingNa, N/A
Gui Add, Text, xs+10 y+10, Delivery Groups?
Gui Add, Radio, x+39 vdeliveryGroupsYes gsubmitChecklist, Yes
Gui Add, Radio, x+5 vdeliveryGroupsNa gsubmitChecklist, N/A
Gui Add, Checkbox, xs+10 y+10 vupdateDeliveryBlock gsubmitChecklist, Update Delivery Block
Gui Add, Text, xs+10 y+10, Order Acceptance
Gui Add, Radio, x+35 vorderAcceptedYes gsubmitChecklist, Yes
Gui Add, Radio, x+5 vorderAcceptedNa gsubmitChecklist, N/A

; ----------- END SAP -----------------

Gui Show,w1000 h700, Order Organizer ;SO# %soNumber%
Gui, Submit, NoHide

submitChecklist:
Gui, Submit, Nohide
return

submitSales:
Gui, Submit, NoHide
gosub, findSales
return

GuiClose:
Gui, destroy
Return

restartScript:
Gosub, WriteIniVariables
if ((!cpq) || (!po))
{
    Reload
} Else
{
    Gosub, SaveToIni
    Reload
} 
return

readtheini:
Gui, Submit, NoHide
if (cpq) && (po)
    gosub, SaveToIniNoGui
FileSelectFile, SelectedFile,r,%myinipath%, Open a file
if (ErrorLevel)
    {
        gosub, restartScript
        return
    }
IniRead, cpq, %SelectedFile%, orderInfo, cpq
GuiControl,, cpq, %cpq%
IniRead, po, %SelectedFile%, orderInfo, po
GuiControl,, po, %po%
IniRead, sot, %SelectedFile%, orderInfo, sot
if % sot == "ERROR"
    sot := 
GuiControl,, sot, %sot%
IniRead, customer, %SelectedFile%, orderInfo, customer
GuiControl,, customer, %customer%
IniRead, salesPerson, %SelectedFile%, orderInfo, salesPerson
GuiControl, Choose, salesPerson, %salesPerson%
IniRead, salesManager, %SelectedFile%, orderInfo, salesManager
GuiControl, Choose, salesManager, %salesManager%
IniRead, salesDirector, %SelectedFile%, orderInfo, salesDirector
GuiControl, Choose, salesDirector, %salesDirector%
IniRead, managerCode, %SelectedFile%, orderInfo, managerCode
GuiControl, ChooseString, managerCode, %managerCode%
IniRead, directorCode, %SelectedFile%, orderInfo, directorCode
GuiControl, ChooseString, directorCode, %directorCode%
IniRead, address, %SelectedFile%, orderInfo, address
GuiControl,, address, %address%
IniRead, contact, %SelectedFile%, orderInfo, contact
GuiControl,, contact, %contact%
IniRead, poValue, %SelectedFile%, orderInfo, poValue
GuiControl,, poValue, %poValue%
IniRead, tax, %SelectedFile%, orderInfo, tax
GuiControl,, tax, %tax%
IniRead, freightCost, %SelectedFile%, orderInfo, freightCost
GuiControl,, freightCost, %freightCost%
IniRead, totalCost, %SelectedFile%, orderInfo, totalCost
GuiControl,, totalCost, %totalCost%
IniRead, system, %SelectedFile%, orderInfo, system
GuiControl,, system, %system%
IniRead, soldTo, %SelectedFile%, orderInfo, soldTo
GuiControl,, soldTo, %soldTo%
IniRead, crd, %SelectedFile%, orderInfo, crd
GuiControl,, crd, %crd%
IniRead, soNumber, %SelectedFile%, orderInfo, soNumber
GuiControl,, soNumber, %soNumber%
IniRead, poDate, %SelectedFile%, orderInfo, poDate
GuiControl,, poDate, %poDate%
IniRead, sapDate, %SelectedFile%, orderInfo, sapDate
GuiControl,, sapDate, %sapDate%
IniRead, endUser, %SelectedFile%, orderInfo, endUser
GuiControl,, endUser, %endUser%
IniRead, phone, %SelectedFile%, orderInfo, phone
GuiControl,, phone, %phone%
IniRead, email, %SelectedFile%, orderInfo, email
GuiControl,, email, %email%

IniRead, endUseEscaped, %SelectedFile%, orderInfo, endUse
; Get back newline separated list.
StringReplace, endUseDeescaped, endUseEscaped, ``n, `n, All
StringReplace, endUseDeescaped, endUseDeescaped, ``r, `r, All
GuiControl,, endUse, %endUseDeescaped%

IniRead, notesEscaped, %SelectedFile%, orderInfo, notes
; Get back newline separated list.
StringReplace, notesDeescaped, notesEscaped, ``n, `n, All
StringReplace, notesDeescaped, notesDeescaped, ``r, `r, All
if % notes == "ERROR"
    notes := 
GuiControl,, notes, %notesDeescaped%

IniRead, software, %SelectedFile%, orderInfo, software
GuiControl,, software, %software%
IniRead, serialNumber, %myinipath%\PO %po%.ini, orderInfo, serialNumber
if (!serialNumber) || (serialNumber == "ERROR")
    GuiControl, Hide, serialNumber
if (serialNumber)
    GuiControl,, serialNumber, %serialNumber%
IniRead, nameCheck, %SelectedFile%, orderInfo, nameCheck
if % nameCheck == "ERROR"
    nameCheck := "Check TENA Name On PO"
GuiControl,, nameCheck, %nameCheck%
IniRead, orderNoticeSent, %SelectedFile%, orderInfo, orderNoticeSent
if % orderNoticeSent == "ERROR"
    orderNoticeSent := "Order Notice Sent"
GuiControl,, orderNoticeSent, %orderNoticeSent%
IniRead, enteredSot, %SelectedFile%, orderInfo, enteredSot
if % enteredSot == "ERROR"
    enteredSot := "Entered In SOT"
GuiControl,, enteredSot, %enteredSot%
IniRead, tandcYes, %SelectedFile%, orderInfo, tandcYes
if % tandcYes == "ERROR"
    tandcYes := "Yes"
GuiControl,, tandcYes, %tandcYes%
IniRead, tandcNa, %SelectedFile%, orderInfo, tandcNa
if % tandcNa == "ERROR"
    tandcNa := "N/A"
GuiControl,, tandcNa, %tandcNa%
IniRead, poAttached, %SelectedFile%, orderInfo, poAttached
if % poAttached == "ERROR"
    poAttached := "PO"
GuiControl,, poAttached, %poAttached%
IniRead, quoteAttached, %SelectedFile%, orderInfo, quoteAttached
if % quoteAttached == "ERROR"
    quoteAttached := "Quote"
GuiControl,, quoteAttached, %quoteAttached%
IniRead, dpsAttached, %SelectedFile%, orderInfo, dpsAttached
if % dpsAttached == "ERROR"
    dpsAttached := "DPS Report(s)"
GuiControl,, dpsAttached, %dpsAttached%
IniRead, orderNoticeAttached, %SelectedFile%, orderInfo, orderNoticeAttached
if % orderNoticeAttached == "ERROR"
    orderNoticeAttached := "Order Notice"
GuiControl,, orderNoticeAttached, %orderNoticeAttached%
IniRead, winYes, %SelectedFile%, orderInfo, winYes
if % winYes == "ERROR"
    winYes := "Yes"
GuiControl,, winYes, %winYes%
IniRead, winNa, %SelectedFile%, orderInfo, winNa
if % winNa == "ERROR"
    winNa := "N/A"
GuiControl,, winNa, %winNa%
IniRead, mergeYes, %SelectedFile%, orderInfo, mergeYes
GuiControl,, mergeYes, %mergeYes%
IniRead, mergeNa, %SelectedFile%, orderInfo, mergeNa
GuiControl,, mergeNa, %mergeNa%
IniRead, checkPrices, %SelectedFile%, orderInfo, checkPrices
GuiControl,, checkPrices, %checkPrices%
IniRead, shippingYes, %SelectedFile%, orderInfo, shippingYes
GuiControl,, shippingYes, %shippingYes%
IniRead, shippingNa, %SelectedFile%, orderInfo, shippingNa
GuiControl,, shippingNa, %shippingNa%
IniRead, higherLevelLinkingYes, %SelectedFile%, orderInfo, higherLevelLinkingYes
GuiControl,, higherLevelLinkingYes, %higherLevelLinkingYes%
IniRead, higherLevelLinkingNa, %SelectedFile%, orderInfo, higherLevelLinkingNa
GuiControl,, higherLevelLinkingNa, %higherLevelLinkingNa%
IniRead, deliveryGroupsYes, %SelectedFile%, orderInfo, deliveryGroupsYes
GuiControl,, deliveryGroupsYes, %deliveryGroupsYes%
IniRead, deliveryGroupsNa, %SelectedFile%, orderInfo, deliveryGroupsNa
GuiControl,, deliveryGroupsNa, %deliveryGroupsNa%
IniRead, updateDeliveryBlock, %SelectedFile%, orderInfo, updateDeliveryBlock
GuiControl,, updateDeliveryBlock, %updateDeliveryBlock%
IniRead, orderAcceptedYes, %SelectedFile%, orderInfo, orderAcceptedYes
GuiControl,, orderAcceptedYes, %orderAcceptedYes%
IniRead, orderAcceptedNa, %SelectedFile%, orderInfo, orderAcceptedNa
GuiControl,, orderAcceptedNa, %orderAcceptedNa%
IniRead, serialYes, %SelectedFile%, orderInfo, serialYes
GuiControl,, serialYes, %serialYes%
IniRead, serialNa, %SelectedFile%, orderInfo, serialNa
GuiControl,, serialNa, %serialNa%
IniRead, endUserYes, %SelectedFile%, orderInfo, endUserYes
GuiControl,, endUserYes, %endUserYes%
IniRead, endUserNa, %SelectedFile%, orderInfo, endUserNa
GuiControl,, endUserNa, %endUserNa%
; GuiControl,, title, Order Organizer - SO# %soNumber%
; WinSetTitle, Order Organizer,,Order Organizer - SO# %soNumber%, Standard Order
return

SaveToIni:
Gui, Submit, NoHide
if (!cpq) || (!po)
{
    MsgBox, Please enter a quote and PO#.
    return
}
IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . ".ini"
IniFilePathWithSo := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . " SO# " . soNumber . ".ini"
if FileExist(IniFilePath) && (soNumber)
{
    gosub, WriteIniVariables
    FileMove, %IniFilePath%, %IniFilePathWithSo% , 1
    IniFilePath = %IniFilePathWithSO% 
    Gosub, SaveBar
    gosub, CheckIfFolderExists
    return  
}
else if FileExist(IniFilePathWithSo)
{
    IniFilePath = %IniFilePathWithSO% 
    Gosub, SaveBar
    gosub, WriteIniVariables
    return
}
else if FileExist(IniFilePath) && (!soNumber)
{
    gosub, WriteIniVariables
    Gosub, SaveBar
    gosub, CheckIfFolderExists
    return
}
else if !FileExist(IniFilePath) && !FileExist(IniFilePathWithSo)
{
    if(soNumber)
    {
        IniFilePath = %IniFilePathWithSo%
    } else {
        IniFilePath = %IniFilePath%
    }
    gosub, WriteIniVariables
    Gosub, SaveBar
    gosub, CheckIfFolderExists
    return
}
return

SaveToIniNoGui:
IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . ".ini"
IniFilePathWithSo := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . " SO# " . soNumber . ".ini"
if FileExist(IniFilePath) && (soNumber)
{
    gosub, WriteIniVariables
    FileMove, %IniFilePath%, %IniFilePathWithSo% , 1
    IniFilePath = %IniFilePathWithSO% 
    gosub, CheckIfFolderExists
    return  
}
else if FileExist(IniFilePathWithSo)
{
    IniFilePath = %IniFilePathWithSO% 
    gosub, WriteIniVariables
    return
}
else if FileExist(IniFilePath) && (!soNumber)
{
    gosub, WriteIniVariables
    gosub, CheckIfFolderExists
    return
}
else if !FileExist(IniFilePath) && !FileExist(IniFilePathWithSo)
{
    if(soNumber)
    {
        IniFilePath = %IniFilePathWithSo%
    } else {
        IniFilePath = %IniFilePath%
    }
    gosub, WriteIniVariables
    gosub, CheckIfFolderExists
    return
}
return

WriteIniVariables:
gui, submit, NoHide
IniWrite, %cpq%, %IniFilePath%, orderInfo, cpq
IniWrite, %po%, %IniFilePath%, orderInfo, po
IniWrite, %sot%, %IniFilePath%, orderInfo, sot
IniWrite, %customer%, %IniFilePath%, orderInfo, customer
IniWrite, %salesPerson%, %IniFilePath%, orderInfo, salesPerson
IniWrite, %salesManager%, %IniFilePath%, orderInfo, salesManager
IniWrite, %managerCode%, %IniFilePath%, orderInfo, managerCode
IniWrite, %salesDirector%, %IniFilePath%, orderInfo, salesDirector
IniWrite, %directorCode%, %IniFilePath%, orderInfo, directorCode
IniWrite, %address%, %IniFilePath%, orderInfo, address
IniWrite, %contact%, %IniFilePath%, orderInfo, contact
IniWrite, %poValue%, %IniFilePath%, orderInfo, poValue
IniWrite, %tax%, %IniFilePath%, orderInfo, tax
IniWrite, %freightCost%, %IniFilePath%, orderInfo, freightCost
IniWrite, %totalCost%, %IniFilePath%, orderInfo, totalCost
IniWrite, %system%, %IniFilePath%, orderInfo, system
IniWrite, %soldTo%, %IniFilePath%, orderInfo, soldTo
IniWrite, %crd%, %IniFilePath%, orderInfo, crd
IniWrite, %soNumber%, %IniFilePath%, orderInfo, soNumber
IniWrite, %poDate%, %IniFilePath%, orderInfo, poDate
IniWrite, %sapDate%, %IniFilePath%, orderInfo, sapDate
IniWrite, %endUser%, %IniFilePath%, orderInfo, endUser
IniWrite, %phone%, %IniFilePath%, orderInfo, phone
IniWrite, %email%, %IniFilePath%, orderInfo, email

    ; Escape all newlines before writing it to ini file.
    StringReplace, endUseEscaped, endUse, `n, ``n, All
    StringReplace, endUseEscaped, endUseEscaped, `r, ``r, All
    IniWrite, %endUseEscaped%, %IniFilePath%, orderInfo, endUse

    ; Escape all newlines before writing it to ini file.
    StringReplace, notesEscaped, notes, `n, ``n, All
    StringReplace, notesEscaped, notesEscaped, `r, ``r, All
    IniWrite, %notesEscaped%, %IniFilePath%, orderInfo, notes

IniWrite, %software%, %IniFilePath%, orderInfo, software
IniWrite, %serialNumber%, %IniFilePath%, orderInfo, serialNumber
if (software == 0)
    IniDelete, %IniFilePath%, orderInfo, serialNumber
IniWrite, %nameCheck%, %IniFilePath%, orderInfo, nameCheck
IniWrite, %orderNoticeSent%, %IniFilePath%, orderInfo, orderNoticeSent
IniWrite, %enteredSot%, %IniFilePath%, orderInfo, enteredSot
IniWrite, %tandcYes%, %IniFilePath%, orderInfo, tandcYes
IniWrite, %tandcNa%, %IniFilePath%, orderInfo, tandcNa
IniWrite, %poAttached%, %IniFilePath%, orderInfo, poAttached
IniWrite, %quoteAttached%, %IniFilePath%, orderInfo, quoteAttached
IniWrite, %dpsAttached%, %IniFilePath%, orderInfo, dpsAttached
IniWrite, %orderNoticeAttached%, %IniFilePath%, orderInfo, orderNoticeAttached
IniWrite, %winYes%, %IniFilePath%, orderInfo, winYes
IniWrite, %winNa%, %IniFilePath%, orderInfo, winNa
IniWrite, %mergeYes%, %IniFilePath%, orderInfo, mergeYes
IniWrite, %mergeNa%, %IniFilePath%, orderInfo, mergeNa
IniWrite, %checkPrices%, %IniFilePath%, orderInfo, checkPrices
IniWrite, %shippingYes%, %IniFilePath%, orderInfo, shippingYes
IniWrite, %shippingNa%, %IniFilePath%, orderInfo, shippingNa
IniWrite, %higherLevelLinkingYes%, %IniFilePath%, orderInfo, higherLevelLinkingYes
IniWrite, %higherLevelLinkingNa%, %IniFilePath%, orderInfo, higherLevelLinkingNa
IniWrite, %deliveryGroupsYes%, %IniFilePath%, orderInfo, deliveryGroupsYes
IniWrite, %deliveryGroupsNa%, %IniFilePath%, orderInfo, deliveryGroupsNa
IniWrite, %updateDeliveryBlock%, %IniFilePath%, orderInfo, updateDeliveryBlock
IniWrite, %orderAcceptedYes%, %IniFilePath%, orderInfo, orderAcceptedYes
IniWrite, %orderAcceptedNa%, %IniFilePath%, orderInfo, orderAcceptedNa
IniWrite, %serialYes%, %IniFilePath%, orderInfo, serialYes
IniWrite, %serialNa%, %IniFilePath%, orderInfo, serialNa
IniWrite, %endUserYes%, %IniFilePath%, orderInfo, endUserYes
IniWrite, %endUserNa%, %IniFilePath%, orderInfo, endUserNa
return

CheckIfFolderExists:
if (RegExMatch(cpq, "(?:^00*)", quoteNumberCpq))
{
    folderPath := A_WorkingDir . "\Order Docs\PO " . po . " " . customer . " - CPQ-" . cpq
} else if (RegExMatch(cpq, "(?:^[2].*)", quoteNumberSap))
{
    folderPath := A_WorkingDir . "\Order Docs\PO %po% %customer% - Quote %cpq%"

} else if (cpq != quoteNumberCpq || cpq != quoteNumbSap)
{
    MsgBox, Invalid Quote
}
if FileExist(folderPath)
    return
if !FileExist(folderPath)
    FileCreateDir, %folderPath%
    Return
return

SaveBar:
WinGetPos x, y, Width, Height, Order Organizer
    x += 350
    y += 50
myRange:=90
Gui, 2: -Caption
Gui, 2:+AlwaysOnTop
Gui, 2: Color, default
Gui, 2: Font, S8 cBlack, Segoe UI Semibold
Gui, 2: Add, Text, x0 y1 w208 h16 +Center, Saving
Gui, 2: Add,Progress, x10 y20 w190 h20 cblue vPro1 Range0-%myRange%,0
;~ Gui,1: Add, Button, x10 w150 h30 glooping, Loop over list
Gui, 2: Show, w210 h50 x847 y313
Gosub, Looping
Return

Looping:
List1:=[1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0]
Pro1:=0
GuiControl,2:,Pro1,% Pro1
Temp:=""
Loop, 20
    {
        Temp.=List1[A_Index]
        Pro1 +=10
        GuiControl,2:,Pro1,% Pro1
        sleep, 20
    }
Gui, 2: Destroy
Return
}

orderInfo()

;******** HOTSTRINGS (TEXT EXPANSION) ********
#c::run calc.exe ; Run calculator
F13::Send, +{F7} ; Next line in item Conditions SAP SOs
;----- Order keyboard shortcuts -----
::zpo::
Send, %po%
return
::zpoz::
Send, PO %po%
return
::zso::
Send, %soNumber%
Return
::zsoz::
Send, SO{#}{Space}%soNumber%
return
::zpq::
Send, %cpq%
return
::zpqz::
Send, CPQ-%cpq%
return
::zval::
Send, %poValue%
return
::zsal::
Send, %salesPerson%
return
::zcust::
Send, %customer%
return
::zcon::
Send, %contact%
return
::zem::
Send, %email%
return
::zsys::
Send, %system%
return
::zenu::
Send, %endUser%
return
::zph::
Send, %phone%
return
::zuse::
Send, %endUse%
return
::zsot::
Send, %sot% ;^{Left}{BackSpace}
return
:O:zcod::Close Out Document
::zcem::Contracts Email - 
::zwin::
Send WIN Form - CPQ-%cpq%
return
::ejim::10246281
;----- End Order keyboard shortcuts -----


;******** END HOTSTRINGS (TEXT EXPANSION) ********

^!v:: ; Show/Hide Order Info GUI
    DetectHiddenWindows, on
    if !WinActive(Order Organizer "ahk_class AutoHotkeyGUI")
    {
        WinActivate, Order Organizer ahk_class AutoHotkeyGUI, 
        ;~ WinWaitActive, Order Info ahk_class AutoHotkeyGUI,
        return
    }
    else if WinActive(Order Organizer "ahk_class AutoHotkeyGUI")
    {
        WinMinimize
        return
    }
return

;/******** MONTIOR COUNT ********/

SysGet, MonitorCount, MonitorCount
SysGet, MonitorPrimary, MonitorPrimary
MsgBox, Monitor Count:`t%MonitorCount%`nPrimary Monitor:`t%MonitorPrimary%
Loop, %MonitorCount%
{
    SysGet, MonitorName, MonitorName, %A_Index%
    SysGet, Monitor, Monitor, %A_Index%
    SysGet, MonitorWorkArea, MonitorWorkArea, %A_Index%
    MsgBox, Monitor:`t#%A_Index%`nName:`t%MonitorName%`nLeft:`t%MonitorLeft% (%MonitorWorkAreaLeft% work)`nTop:`t%MonitorTop% (%MonitorWorkAreaTop% work)`nRight:`t%MonitorRight% (%MonitorWorkAreaRight% work)`nBottom:`t%MonitorBottom% (%MonitorWorkAreaBottom% work)
}

X := 250, Y := 250 ; Starting position for the Gui on your main monitor
CoordMode, Mouse, Screen
MouseGetPos, MX, MY
If (MX > A_ScreenWidth)
    X += A_ScreenWidth
Gui, Show, x%X% y%Y% w300 h300
return

kellerDropDown:
    GuiControl, Choose, salesManager, Anjou Keller
    GuiControl, ChooseString, managerCode, 202375
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, N/A
    GuiControl, ChooseString, directorCode, N/A
    Gui, Submit, NoHide
return

mccormackDropDown:
    GuiControl, Choose, salesManager, Doug McCormack
    GuiControl, ChooseString, managerCode, 1261
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, N/A
    GuiControl, ChooseString, directorCode, N/A
    Gui, Submit, NoHide
return

hewittDropDown:
    GuiControl, Choose, salesManager, Joe Hewitt
    GuiControl, ChooseString, managerCode, 98866
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, N/A
    GuiControl, ChooseString, directorCode, N/A
    Gui, Submit, NoHide
return

foelsDropDown:
    GuiControl, Choose, salesManager, Natalie Foels
    GuiControl, ChooseString, managerCode, 96715
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, N/A
    GuiControl, ChooseString, directorCode, N/A
    Gui, Submit, NoHide
return

secondDropDown:
    GuiControl, Choose, salesManager, Tonya Second
    GuiControl, ChooseString, managerCode, 95410
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, Maroun El Khoury
    GuiControl, ChooseString, directorCode, 1076
    Gui, Submit, NoHide
return

mcfaddenDropDown:
    GuiControl, Choose, salesManager, Joe McFadden
    GuiControl, ChooseString, managerCode, 202610
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, N/A
    GuiControl, ChooseString, directorCode, N/A
    Gui, Submit, NoHide
return

butlerDropDown:
    GuiControl, Choose, salesManager, John Butler
    GuiControl, ChooseString, managerCode, 1026
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, Maroun El Khoury
    GuiControl, ChooseString, directorCode, 1076
    Gui, Submit, NoHide
return

kleinDropDown:
    GuiControl, Choose, salesManager, Richard Klein
    GuiControl, ChooseString, managerCode, 1042
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, Maroun El Khoury
    GuiControl, ChooseString, directorCode, 1076
    Gui, Submit, NoHide
return

nadjieDropDown:
    GuiControl, Choose, salesManager, Zee Nadjie
    GuiControl, ChooseString, managerCode, 96695
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, N/A
    GuiControl, ChooseString, directorCode, N/A
    Gui, Submit, NoHide
return

chenDropDown:
    GuiControl, Choose, salesManager, Ray Chen
    GuiControl, ChooseString, managerCode, 202756
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, Jimmy Yuk
    GuiControl, ChooseString, directorCode, 202611
    Gui, Submit, NoHide
return

porchDropDown:
    GuiControl, Choose, salesManager, Randy Porch
    GuiControl, ChooseString, managerCode, 1041
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, Maroun El Khoury
    GuiControl, ChooseString, directorCode, 1076
    Gui, Submit, NoHide
return

findSales:
;=============== KELLER ===================
if salesPerson = Julie Sawicki
    gosub, kellerDropDown
if salesPerson = Justin Carder 
    gosub, kellerDropDown
if salesPerson = Brent Boyle
    gosub, kellerDropDown
if salesPerson = Luke Marty
    gosub, kellerDropDown
if salesPerson = Aeron Avakian 
    gosub, kellerDropDown
if salesPerson = Jon Needels
    gosub, kellerDropDown
if salesPerson = Brian Dowe
    gosub, kellerDropDown
if salesPerson = Lisa Kasper
    gosub, kellerDropDown
if salesPerson = John Bailey
    gosub, kellerDropDown
if salesPerson = Aeron Avakian
    gosub, kellerDropDown
if salesPerson = Luke Marty
    gosub, kellerDropDown
if salesPerson = Brandon Markle
    gosub, kellerDropDown
;=============== END KELLER ===================

;============ NADJIE ================
if salesPerson = Jawad Pashmi
    gosub, nadjieDropDown
if salesPerson = Navette Shirakawa
    gosub, nadjieDropDown
if salesPerson = Alicia Arias
    gosub, nadjieDropDown
if salesPerson = Gabriel Mendez
    gosub, nadjieDropDown
if salesPerson = Rhonda Oesterle
    gosub, nadjieDropDown
if salesPerson = Alexander James
    gosub, nadjieDropDown
if salesPerson = Shijun Sheng
    gosub, nadjieDropDown
if salesPerson = Brian Luckenbill
    gosub, nadjieDropDown
if salesPerson = Amy Allgower
    gosub, nadjieDropDown
if salesPerson = Michael Burnett
    gosub, nadjieDropDown
;============ END NADJIE ================

;========== DOUG MCCORMACK =================
if salesPerson = Mark Krigbaum
    gosub, mccormackDropDown
if salesPerson = Samantha Stikeleather
    gosub, mccormackDropDown
if salesPerson = Jeff Weller
    gosub, mccormackDropDown
if salesPerson = Jerry Holycross
    gosub, mccormackDropDown
if salesPerson = Dan Ciminelli
    gosub, mccormackDropDown
if salesPerson = Theresa Borio
    gosub, mccormackDropDown
if salesPerson = Cynthia Spittler
    gosub, mccormackDropDown
if salesPerson = Fred Simpson
    gosub, mccormackDropDown
if salesPerson = Gwyn Trojan
    gosub, mccormackDropDown
if salesPerson = Nick Hubbard
    gosub, mccormackDropDown
if salesPerson = Kristen Luttner
    gosub, mccormackDropDown
;========== END DOUG MCCORMACK =================

;=========== HEWITT ============
if salesPerson = Douglas Sears
    gosub, hewittDropDown
if salesPerson = Melissa Chandler
    gosub, hewittDropDown
if salesPerson = Hillary Tennant
    gosub, hewittDropDown
if salesPerson = Don Rathbauer
    gosub, hewittDropDown
if salesPerson = Tucker Lincoln
    gosub, hewittDropDown
if salesPerson = Joel Stradtner
    gosub, hewittDropDown
if salesPerson = Stephanie Koczur
    gosub, hewittDropDown
if salesPerson = Mike Hughes
    gosub, hewittDropDown
;=========== END HEWITT ============

;============ FOELS ================
if salesPerson = Larry Bellan
    gosub, foelsDropDown
if salesPerson = Kevin Clodfelter
    gosub, foelsDropDown
if salesPerson = Brian Thompson
    gosub, foelsDropDown
if salesPerson = Rashila Patel
    gosub, foelsDropDown
if salesPerson = Bill Balsanek
    gosub, foelsDropDown
if salesPerson = Chuck Costanza
    gosub, foelsDropDown
if salesPerson = Karl Kastner
    gosub, foelsDropDown
if salesPerson = Bob Riggs
    gosub, foelsDropDown
if salesPerson = Drew Smillie
    gosub, foelsDropDown
if salesPerson = Crystal Flowers
    gosub, foelsDropDown
;============ END FOELS ================

;========= SECOND ==============
if salesPerson = Helen Sun
    gosub, secondDropDown
if salesPerson = Steven Danielson
    gosub, secondDropDown
if salesPerson = Dominique Figueroa
    gosub, secondDropDown
if salesPerson = Jonathan McNally
    gosub, secondDropDown
if salesPerson = Yan Chen
    gosub, secondDropDown
if salesPerson = Katianna Pihakari
    gosub, secondDropDown
if salesPerson = Timothy Johnson
    gosub, secondDropDown
;========= END SECOND ==============

;========= MCFADDEN ==============
if salesPerson = May Chou
    gosub, mcfaddenDropDown
if salesPerson = Steve Boyanoski
    gosub, mcfaddenDropDown
if salesPerson = Mark Woodworth
    gosub, mcfaddenDropDown
if salesPerson = Murray Fryman
    gosub, mcfaddenDropDown
if salesPerson = Lorraine Foglio
    gosub, mcfaddenDropDown
if salesPerson = Lauren Fischer
    gosub, mcfaddenDropDown
if salesPerson = Douglas McDowell
    gosub, mcfaddenDropDown
;========= END MCFADDEN ==============

;=========== BUTLER ==========
if salesPerson = Andrew Clark
    gosub, butlerDropDown
if salesPerson = Giovanni Pallante
    gosub, butlerDropDown
if salesPerson = David Kage
    gosub, butlerDropDown
if salesPerson = David Scott
    gosub, butlerDropDown
if salesPerson = Susan Gelman
    gosub, butlerDropDown
if salesPerson = Cari Randles
    gosub, butlerDropDown
if salesPerson = Sean Bennett
    gosub, butlerDropDown
;=========== END BUTLER ==========

;=========== KLEIN ==========
if salesPerson = Susan Bird
    gosub, kleinDropDown
if salesPerson = Jerry Pappas
    gosub, kleinDropDown
if salesPerson = Jie Qian
    gosub, kleinDropDown
if salesPerson = Joe Bernholz
    gosub, kleinDropDown
if salesPerson = Yuriy Dunayevskiy
    gosub, kleinDropDown
if salesPerson = Nelson Huang
    gosub, kleinDropDown
;=========== END KLEIN ==========

;=========== CHEN ==========
if salesPerson = Haris Dzaferbegovic
    gosub, chenDropDown
if salesPerson = Donna Zwirner
    gosub, mccormackDropDown
;=========== END CHEN ==========

;=========== PORCH ==========
if salesPerson = Todd Stoner
    gosub, porchDropDown
if salesPerson = Jonathan Ferguson
    gosub, porchDropDown
if salesPerson = Nick Duczak
    gosub, porchDropDown
if salesPerson = Gerald Koncar
    gosub, porchDropDown
;=========== END PORCH ==========

;======== BLANKS =======
if salesPerson =
{
    GuiControl, Choose, salesManager, |1 
    GuiControl, Choose, managerCode, |1
    Gui, Submit, NoHide
    GuiControl, Choose, salesDirector, |1
    GuiControl, Choose, directorCode, |1
    Gui, Submit, NoHide
}
return

