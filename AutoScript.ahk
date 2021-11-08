#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir, C:\Users\%A_UserName%\OneDrive - Thermo Fisher Scientific\Documents ; Ensures a consistent starting directory.
#SingleInstance Force

salesPeople := "|Justin Carder|Robin Sutka|Fred Simpson|Rhonda Oesterle|Mitch Lazaro|Tucker Lincoln|Jawad Pashmi|Julie Sawicki|Mike Hughes|Steve Boyanoski"
. "|Bob Riggs|Chuck Costanza|Navette Shirakawa|Stephanie Koczur|Mark Krigbaum|Jon Needels|Bill Balsanek|Brent Boyle|Andrew Clark|Kevin Clodfelter|Gabriel Mendez"
. "|Karl Kastner|Michael Burnett|Jerry Pappas|Nick Duczak|Steven Danielson|Nick Hubbard|Samantha Stikeleather|Drew Smillie|Jeff Weller|Jerry Holycross"
. "|Theresa Borio|Dan Ciminelli|Cynthia Spittler|Gwyn Trojan|Joel Stradtner|Don Rathbauer|Hillary Tennant|Melissa Chandler|Douglas Sears|Rashila Patel|Brian Thompson"
. "|Larry Bellan|Donna Zwirner|Kristen Luttner|Helen Sun|May Chou|Haris Dzaferbegovic|Brian Dowe|Mark Woodworth|Susan Bird|Giovanni Pallante|Alicia Arias"
. "|Dominique Figueroa|Jonathan McNally|Murray Fryman|Yan Chen|Jie Qian|Joe Bernholz|David Kage|David Scott|Todd Stoner|John Bailey|Katianna Pihakari|Jonathan Ferguson"
. "|Aeron Avakian|Luke Marty|Alexander James|Timothy Johnson|Yuriy Dunayevskiy|Susan Gelman|Cari Randles|Shijun Sheng|Sean Bennett|Nelson Huang|Lorraine Foglio|Gerald Koncar"
. "|Lauren Fischer|Brian Luckenbill|Amy Allgower|Brandon Markle|Crystal Flowers|"

salesManagers := "|Anjou Keller|Joe Hewitt|Zee Nadjie|Doug McCormack|Natalie Foels|Tonya Second|Lou Gavino|Christopher Crafts|Joe McFadden|John Butler|Richard Klein|Ray Chen|Randy Porch"
salesDirectors := "|Joann Purkerson|Maroun El Khoury|Jimmy Yuk|N/A"
salesCodes := "|202375|96715|1261|98866|96695|96654|202625|202006|1076|95410|202610|1026|1042|202756|202611|1041|N/A"

; Set Order Organizer Path
myinipath := A_WorkingDir 

; Include Icon
FileInstall, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Desktop\Auto Hot Key Scripts\list_check_checklist_checkmark_icon_181579.ico, A_WorkingDir, 1

SetTitleMatchMode, 2

;/******** GUI START ********\
orderInfo(){
    global
Gui, destroy
Gui -DPIScale
Gui, Font
Gui Font, s12 w600 Italic cBlack, Tahoma
Gui Add, Text, hWndhTxtOrderDetails23 x15 y-2 w300 h33 +0x200 +Left, Order Organizer - SO# %soNumber%
Gui, Font
Gui, Color, 79b8d1
Gui, Font, S9, Segoe UI Semibold
Gui, Add, Button, xm+525 ym+10 w70 greadtheini, O&pen
Gui, Add, Button, x+25 w70 gSaveToIni, &Save
Gui, Add, Button, x+25 w150 grestartScript, &New PO or Reload
;******** Placeholder for search ********
; Gui Add, Edit, x+25 y22 w175 h20, Search
Gui Add, Tab3, xm ym+30, Order Info|Checklist
Gui Tab, 1
Gui Add, Text,, CPQ:
Gui Add, Edit, vcpq, %cpq% 
Gui Add, Text,, PO:
Gui Add, Edit, vpo, %po% 
Gui Add, Text,, SOT Line#
Gui Add, Edit, vsot, %sot%
Gui Add, Text,, Customer:
Gui Add, Edit, vcustomer, %customer% 
Gui Add, Text,, Customer / DPS Contact:
Gui Add, Edit, vcontact, %contact% 
Gui Add, Text,, Customer / DPS Address:
Gui Add, Edit, vaddress, %address% 
Gui Add, Text,, Sold To Account:
Gui Add, Edit, vsoldTo, %soldTo%
Gui, Add, GroupBox, x+12.5 y100 w1 h347 ; vertical line
Gui Add, Text, x+12.5 y72, System:
Gui Add, Edit, vsystem, %system% 
Gui, Add, Text,, Salesperson:
Gui, Add, DropDownList, +Sort vsalesPerson gsubmitSales, % salesPeople
Gui, Add, Text,, Sales Manager:
Gui, Add, DDL, Disabled vsalesManager, % salesManagers
Gui Add, Text,, Sales Manager Code:
Gui Add, DropDownList, ReadOnly vmanagerCode, % salesCodes
Gui, Add, Text,, Sales Director:
Gui, Add, DropDownList, Disabled vsalesDirector, % salesDirectors
Gui Add, Text,, Sales Director Code:
Gui, Add, DropDownList, ReadOnly vdirectorCode, % salesCodes
Gui Add, CheckBox, x220 y430 vsoftware gdongle, Software?
Gui, Add, Text, x190 y435.5 Hidden vserialNumberText, S/N:  ;x187.5 y460.5 h10, S/N: ;x187.5 y437.5
Gui Add, Edit, x225 y430 w100 Hidden vserialNumber, ; y432.5
Gui, Add, GroupBox, x+12.5 y100 w1 h347 ; vertical line
Gui Add, Text, x+12.5 y70, CRD:
Gui Add, DateTime, w135 vcrd, %crd%
Gui Add, Text,, PO Date:
Gui Add, DateTime, w135 vpoDate, %poDate% 
Gui Add, Text,, SAP Date:
Gui Add, DateTime, w135 vsapDate, %sapDate%
Gui Add, Text,, PO Value:
Gui Add, Edit, vpoValue, %poValue%
Gui Add, Text,, Tax:
Gui Add, Edit, vtax, %tax%
Gui Add, Text,, Freight Cost:
Gui Add, Edit, vfreightCost, %freightCost% 
Gui Add, Text,, Total:
Gui Add, Edit, vtotalCost, %totalCost% 
Gui, Add, GroupBox, x+12.5 y100 w1 h347 ; vertical line
Gui Add, Text, x+12.5 y95 +center, END USER INFO:
Gui Add, Text,, End User:
Gui Add, Edit, vendUser
Gui Add, Text,, Phone:
Gui Add, Edit, vphone
Gui Add, Text,, Email:
Gui Add, Edit, vemail
Gui Add, Text,, End Use:
Gui Add, Edit, w135 h78 vendUse
Gui Add, Text,, SO#
Gui Add, Edit, vsoNumber, %soNumber% 
Gui, Add, GroupBox, x+12.5 y100 w1 h347 ; vertical line
Gui Add, Text, x+12.5 y70, Notes:
Gui Add, Edit, w215 h120 vnotes, %notes%

;======== KEYBOARD SHORTCUTS ========
Gui Add, Listview, y+10 w215 h235 R13 grid ReadOnly, Value (Keyboard Shortcut)
LV_ModifyCol(1,190)
;~ LV_ModifyCol(2, 115)
LV_Add(Col1, "CPQ (zpq)") ;"zpq")
;~ Gui Add, Text,, CPQ - zpq
LV_Add(Col1, "PO# (zpo)") ;,"zpo")
;~ Gui Add, Text, y+5, PO# - zpo
LV_Add(Col1, "SO# (zso)") ;,"zso")
;~ Gui Add, Text, y+5, SO# - zso
LV_Add(Col1, "SOT Line# (zsot)") ;,"zsot")
;~ Gui Add, Text, y+5, SOT Line# - zsot
LV_Add(Col1, "Customer (zcust)") ;,"zcust")
;~ Gui Add, Text, y+5, Customer - zcust
LV_Add(Col1, "PO Value (zval)") ;,"zval")
;~ Gui Add, Text, y+5, PO Value - zval
LV_Add(Col1, "Salesperson (zsal)") ;,"zsal")
;~ Gui Add, Text, y+5, Salesperson - zsal
LV_Add(Col1, "Cust Contact (zcon)") ;,"zcon")
;~ Gui Add, Text, x775 y225, Customer Contact -
;~ Gui Add, Text, y+5, zcon
LV_Add(Col1, "Cust Email (zem)")
;~ Gui Add, Text, y+5, Customer Email - 
;~ Gui Add, Text, y+5, zem
LV_Add(Col1, "System (zsys)")
;~ Gui Add, Text, y+5, System - zsys
LV_Add(Col1, "End User (zenu)")
;~ Gui Add, Text, y+5, End User - zenu
LV_Add(Col1, "End User Phone (zph)")
;~ Gui Add, Text, y+5, Phone - zph
LV_Add(Col1, "End Use (zuse)")
;~ Gui Add, Text, y+5, End Use - zuse
;======== END KEYBOARD SHORTCUTS ========

;******** CHECKLIST GUI ********
Gui, Tab, 2

;===== PRE-SAP =======
Gui, Add, GroupBox,Section h120 w215, Pre-SAP
Gui, Add, Checkbox, xs+10 ys+20 gsubmitChecklist vnameCheck, Check TENA Name On PO
Gui, Add, Checkbox, gsubmitChecklist vorderNoticeSent, Order Notice Sent
Gui, Add, Checkbox, gsubmitChecklist venteredSot, Entered In SOT
Gui, Add, Text, tcs y+10, T&&Cs?
Gui, Add, Radio, x+5 gsubmitChecklist vtandcYes, Yes
Gui, Add, Radio, x+5 gsubmitChecklist vtandcNa, N/A
;===== END PRE-SAP =======

;===== SAP ATTACHMENTS =======
Gui, Add, GroupBox,Section xm+15 y+25 h130 w215, SAP Attachments:
Gui, Add, Checkbox,xs+10 ys+25 gsubmitChecklist vpoAttached, PO
Gui, Add, Checkbox, x+61 gsubmitChecklist vquoteAttached, Quote
Gui, Add, Checkbox, xm+25 y+10 gsubmitChecklist vdpsAttached, DPS Report(s)
Gui, Add, Checkbox, x+5 gsubmitChecklist vorderNoticeAttached, Order Notice
Gui, Add, Text, xm+25 y+15, WIN Form?
Gui, Add, Radio, x+40 gsubmitChecklist vwinYes, Yes
Gui, Add, Radio, x+5 gsubmitChecklist vwinNa, N/A
Gui, Add, Text, xm+25 y+10, Merge Report?
Gui, Add, Radio, x+24 gsubmitChecklist vmergeYes, Yes
Gui, Add, Radio, x+5 gsubmitChecklist vmergeNa, N/A
;===== END SAP ATTACHMENTS =======

;===== PRE ACCEPTANCE ========
Gui, Add, GroupBox,Section xm+260 ym+63 h225 w235, Pre-Acceptance:
Gui, Add, Checkbox, xs+10 ys+25 gsubmitChecklist vcheckPrices, Check Prices
Gui, Add, Text,, Add Shipping?
Gui, Add, Radio, x+53 gsubmitChecklist vshippingYes, Yes
Gui, Add, Radio, x+5 gsubmitChecklist vshippingNa, N/A
Gui, Add, Text, xm+270 y+10, Higher Level Linking?
Gui, Add, Radio, x+18 gsubmitChecklist vhigherLevelLinkingYes, Yes
Gui, Add, Radio, x+5 vhigherLevelLinkingNa, N/A
Gui, Add, Text,xm+270 y+10, Delivery Groups?
Gui, Add, Radio, x+42.5 vdeliveryGroupsYes gsubmitChecklist, Yes
Gui, Add, Radio, x+5 vdeliveryGroupsNa gsubmitChecklist, N/A
Gui, Add, Checkbox, xm+270 y+10 vupdateDeliveryBlock gsubmitChecklist, Update Delivery Block
Gui, Add, Text,, Order Acceptance
Gui, Add, Radio, x+35 vorderAcceptedYes gsubmitChecklist, Yes
Gui, Add, Radio, x+5 vorderAcceptedNa gsubmitChecklist, N/A
Gui, Add, Text, xm+270 y+10, Serial/Dongle Number?
Gui, Add, Radio, x+5 gsubmitChecklist vserialYes, Yes
Gui, Add, Radio, x+5 gsubmitChecklist vserialNa, N/A
Gui, Add, Text, xm+270 y+10, End User Info?
Gui, Add, Radio, x+54 gsubmitChecklist vendUserYes, Yes
Gui, Add, Radio, x+5 gsubmitChecklist vendUserNa, N/A
;===== END PRE ACCEPTANCE ========
;~ Gui, add, Text, x60 y400 , Order Progress
;~ Gui, Add, Progress, w800 h25, 25
;******** END CHECKLIST GUI ********

;======== KEYBOARD SHORTCUTS ========
Gui Add, Listview, xm+525 ym+70 w215 h275 R13 grid ReadOnly, Value (Keyboard Shortcut)
LV_ModifyCol(1,190)
;~ LV_ModifyCol(2, 115)
LV_Add(Col1, "CPQ (zpq)") ;"zpq")
;~ Gui Add, Text,, CPQ - zpq
LV_Add(Col1, "PO# (zpo)") ;,"zpo")
;~ Gui Add, Text, y+5, PO# - zpo
LV_Add(Col1, "SO# (zso)") ;,"zso")
;~ Gui Add, Text, y+5, SO# - zso
LV_Add(Col1, "SOT Line# (zsot)") ;,"zsot")
;~ Gui Add, Text, y+5, SOT Line# - zsot
LV_Add(Col1, "Customer (zcust)") ;,"zcust")
;~ Gui Add, Text, y+5, Customer - zcust
LV_Add(Col1, "PO Value (zval)") ;,"zval")
;~ Gui Add, Text, y+5, PO Value - zval
LV_Add(Col1, "Salesperson (zsal)") ;,"zsal")
;~ Gui Add, Text, y+5, Salesperson - zsal
LV_Add(Col1, "Cust Contact (zcon)") ;,"zcon")
;~ Gui Add, Text, x775 y225, Customer Contact -
;~ Gui Add, Text, y+5, zcon
LV_Add(Col1, "Cust Email (zem)")
;~ Gui Add, Text, y+5, Customer Email - 
;~ Gui Add, Text, y+5, zem
LV_Add(Col1, "System (zsys)")
;~ Gui Add, Text, y+5, System - zsys
LV_Add(Col1, "End User (zenu)")
;~ Gui Add, Text, y+5, End User - zenu
LV_Add(Col1, "End User Phone (zph)")
;~ Gui Add, Text, y+5, Phone - zph
LV_Add(Col1, "End Use (zuse)")
;~ Gui Add, Text, y+5, End Use - zuse
;======== END KEYBOARD SHORTCUTS ========

Gui Show,w920 h485, Order Organizer SO# %soNumber%
Gui, Submit, NoHide

submitChecklist:
Gui, Submit, Nohide
return

submitSales:
Gui, Submit, NoHide
gosub, findSales
return

dongle:
Gui, Submit, NoHide
if (software == 0)
	{
		GuiControl, Hide, serialNumber
		GuiControl, Hide, serialNumberText
		GuiControl, Move, software, y368
	}
if (software == 1)
	{
		GuiControl, Move, software, y345
		GuiControl, Show, serialNumber
		GuiControl, Show, serialNumberText
	}
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
if (cpq) || (po)
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
IniRead, endUse, %SelectedFile%, orderInfo, endUse
GuiControl,, endUse, %endUse%
IniRead, notes, %SelectedFile%, orderInfo, notes
if % notes == "ERROR"
	notes := 
GuiControl,, notes, %notes%
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
return

SaveToIni:
Gosub, WriteIniVariables
if (!cpq) || (!po)
{
	MsgBox, Please enter a quote and PO#.
	return
}
IniFilePath := myinipath . "\PO " . po . " " . customer . ".ini"
IniFilePathWithSo := myinipath . "\PO " . po . " " . customer . " SO# " . soNumber . ".ini"
if FileExist(IniFilePath) && (soNumber)
{
	FileMove, %IniFilePath%, %IniFilePathWithSo% , 1
	IniFilePath = %IniFilePathWithSO% 
	gosub, WriteIniVariables
	; gosub, SaveBar
	gosub, CheckIfFolderExists
	return
}
else if FileExist(IniFilePathWithSo)
{
	IniFilePath = %IniFilePathWithSO% 
	gosub, WriteIniVariables
	; gosub, SaveBar
	return
}
else if FileExist(IniFilePath) && (!soNumber)
{
	gosub, WriteIniVariables
	; gosub, SaveBar
	gosub, CheckIfFolderExists
	return
}
else if !FileExist(IniFilePath) && !FileExist(IniFilePathWithSo)
{
	gosub, WriteIniVariables
	; gosub, SaveBar
	gosub, CheckIfFolderExists
	return
}
return

SaveToIniNoGui:
IniFilePath := myinipath . "\PO " . po . " " . customer . ".ini"
IniFilePathWithSo := myinipath . "\PO " . po . " " . customer . " SO# " . soNumber . ".ini"
if FileExist(IniFilePath) && (soNumber)
{
	FileMove, %IniFilePath%, %IniFilePathWithSo% , 1
	IniFilePath = %IniFilePathWithSO% 
	gosub, WriteIniVariables
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
IniWrite, %endUse%, %IniFilePath%, orderInfo, endUse
IniWrite, %notes%, %IniFilePath%, orderInfo, notes
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
folderPath = C:\Users\%A_UserName%\OneDrive - Thermo Fisher Scientific\Documents\PO %po% %customer% - CPQ-%cpq%
} else if (RegExMatch(cpq, "(?:^[2].*)", quoteNumberSap))
{
folderPath = C:\Users\%A_UserName%\OneDrive - Thermo Fisher Scientific\Documents\PO %po% %customer% - Quote %cpq%
} else if (cpq != quoteNumberCpq || cpq != quoteNumbSap)
{
MsgBox, Invalid Quote
}
if FileExist(folderPath)
	return
if !FileExist(folderPath)
	FileCreateDir, %folderPath%
	run, C:\Users\%A_UserName%\OneDrive - Thermo Fisher Scientific\Documents\
    Return
return

; SaveBar:
; WinGetPos x, y, Width, Height, Order Info
;     x := x + 350
; 	y := y + 30
; myRange:=90
; Gui, 2: -Caption
; Gui, 2:+AlwaysOnTop
; Gui, 2: Color, default
; Gui, 2: Font, S8 cBlack, Segoe UI Semibold
; Gui, 2: Add, Text, x0 y1 w208 h16 +Center, Saving
; Gui, 2: Add,Progress, x10 y20 w190 h20 cblue vPro1 Range0-%myRange%,0
; ;~ Gui,1: Add, Button, x10 w150 h30 glooping, Loop over list
; Gui, 2: Show, x%X% y%Y% w210 h50 
; Gosub, Looping
; Return

; Looping:
; List1:=[1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0]
; Pro1:=0
; GuiControl,2:,Pro1,% Pro1
; Temp:=""
; Loop, 20
; 	{
; 		Temp.=List1[A_Index]
; 		Pro1 +=10
; 		GuiControl,2:,Pro1,% Pro1
; 		sleep, 20
; 	}
; Gui, 2: Destroy
; Return
}

orderInfo()

;******** HOTSTRINGS (TEXT EXPANSION) ********
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

;******** END HOTSTRINGS (TEXT EXPANSION) ********

^!v:: ; Show/Hide Order Info GUI
DetectHiddenWindows, on
if !WinActive(Order Info "ahk_class AutoHotkeyGUI")
{
	WinActivate, Order Info ahk_class AutoHotkeyGUI, 
	;~ WinWaitActive, Order Info ahk_class AutoHotkeyGUI,
	return
}
else if WinActive(Order Info "ahk_class AutoHotkeyGUI")
{
	WinMinimize
	return
}
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