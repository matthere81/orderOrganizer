﻿#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
;~ #Include WatchFolder.ahk
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs ; Ensures a consistent starting directory.
#SingleInstance Force
#InstallKeybdHook
#KeyHistory 50
#include <Vis2>
#include <UIA_Interface>
#include <UIA_Browser>
; #include OrderOrganizerFunctions.ahk

salesDirectors := "|Denise Schwartz|Joann Purkerson|Maroun El Khoury|Jimmy Yuk|Sylveer Bergs|N/A"

salesCodes := "|201020|202375|96715|1261|98866|96695|96654|202625|202006|1076|95410|202610|1026|1042|202756|202611|1041|200320|203185|1416|203915|N/A"

myinipath = C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB

I_Icon = C:\Users\%A_UserName%\OneDrive - Thermo Fisher Scientific\Desktop\Auto Hot Key Scripts\list_check_checklist_checkmark_icon_181579.ico
IfExist, %I_Icon%
	Menu, Tray, Icon, %I_Icon%

Menu, FileMenu, Add

SetTitleMatchMode, 2

; Initialize a global variable to keep track of the visibility
global checkListVisible := true
global guiControls := ["nameCheck", "quoteNumberMatch", "paymentTerms", "priceMatch", "bothAddresses", "pdfQuote", "arrangeLines", "soldToIdCheck", "orderTypeCheck", "poInfoCheck", "generateDps", "orderNoticeSent", "enteredSot", "tandcYes", "tandcNa", "poAttached", "quoteAttached", "dpsAttached", "orderNoticeAttached", "winYes", "winNa", "mergeYes", "mergeNa", "crdDateAdded", "firstDate", "incoterms", "volts", "shipper", "verifyIncoterms", "managerCodeCheck", "directorCodeCheck", "billtoCheck", "shiptoCheck", "endUserCheck", "endUserNA", "contactPersonCheck", "textsContactCheck", "serialYes", "serialNa", "endUserCopyBack", "endUserCopyBackNa", "finalTotal", "shippingYes", "shippingNa", "higherLevelLinkingYes", "higherLevelLinkingNa", "deliveryGroupsYes", "deliveryGroupsNa", "updateDeliveryBlock", "orderAcceptedYes", "orderAcceptedNa"]


orderInfo(){
global
;/******** GUI START ********\
Gui destroy
Gui Font
Gui Font, s12 w600 Italic cBlack, Tahoma
Gui Add, Text, hWndhTxtOrderOrganizer23 x15 y-2 w300 h33 +0x200 +Left, Order Organizer ; - SO# %soNumber%
Gui Font
Gui Color, 79b8d1
Gui Font, S9, Segoe UI Semibold
LVArray := []

Gui Add, Edit, x235 y19 w125 h27.5 vSearchTerm gSearch, ; LV Search
Gui Add, ListView, grid r5 w400 y50 x250 vLV gMyListView, ORDERS:


Loop, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB\*.*
{
LV_Add("", A_LoopFileName)
LVArray.Push(A_LoopFileName)
}
GuiControl, hide, LV


; Gui Add, DDL, w400 y50 x250 vLV gMyListView, ORDERS:
; Loop, % LVArray.MaxIndex()
; 	DDLString .= "|" LVArray[A_Index]
Gui Add, Button, xm+370 ym+10 w70 greadtheini, O&pen
Gui Add, Button, x+25 w70 gSaveToIni, &Save
Gui Add, Button, x+25 w150 grestartScript, &New PO or Reload
Gui Add, Button, x+25 w150 gClearFields, &Clear Fields
Gui Add, Tab3, xm ym+30, Order Info|Checklist
Gui Tab, 1
Gui Add, Text, Section, CPQ:
Gui Add, Edit, vcpq, %cpq% 
Gui Add, Text,, PO:
Gui Add, Edit, vpo, %po%
Gui Add, Text,, SOT Line#
Gui Add, Edit, vsot, %sot%
Gui Add, Text,, Customer:
Gui Add, Edit, vcustomer, %customer% 
Gui Add, Text,, Customer / DPS Contact:
Gui Add, Edit, vcontact, %contact% 
Gui Add, Text,, Intercompany? - Entity
Gui Add, Edit, vaddress, %address% 
Gui Add, Text,, Sold To Account:
Gui Add, Edit, vsoldTo, %soldTo%
Gui Add, Text,, Payment Terms
Gui Add, Edit, vterms, %terms% 
Gui Add, GroupBox, x+12.5 y100 w1 h400 ; vertical line
Gui Add, Text, x+12.5 y72 Section, System:
Gui Add, Edit, vsystem, %system% 
Gui Add, Text,, Salesperson:
; Gui Add, DropDownList, +Sort vsalesPerson gsubmitSales, % salesPeople
Gui Add, Edit, vsalesPerson, %salesPerson%
Gui Add, Text,, Sales Manager:
; Gui Add, DDL, Disabled vsalesManager, % salesManagers
Gui Add, Edit, vsalesManager gsubmitSales, %salesManager%
Gui Add, Text,, Sales Manager Code:
Gui Add, DropDownList, ReadOnly vmanagerCode, % salesCodes
Gui Add, Text,, Sales Director:
Gui Add, DropDownList, Disabled vsalesDirector, % salesDirectors
Gui Add, Text,, Sales Director Code:
Gui Add, DropDownList, ReadOnly vdirectorCode, % salesCodes
Gui Add, CheckBox, x220 y430 vsoftware gdongle, Software?
Gui Add, Text, x190 y435.5 Hidden vserialNumberText, S/N:  ;x187.5 y460.5 h10, S/N: ;x187.5 y437.5
Gui Add, Edit, x225 y430 w100 Hidden vserialNumber, ; y432.5
Gui Add, GroupBox, x+12.5 y100 w1 h400 ; vertical line
Gui Add, Text, x+12.5 y70 Section, CRD:
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
Gui Add, Text,, Surcharge:
Gui Add, Edit, w135 vsurcharge, %surcharge%
Gui Add, Text,, Total:
Gui Add, Edit, vtotalCost, %totalCost% 
Gui Add, GroupBox, x+12.5 y100 w1 h400 ; vertical line
Gui Add, Text, x+12.5 y95 +center Section, END USER INFO:
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
Gui Add, GroupBox, x+12.5 y100 w1 h400 ; vertical line
Gui Add, Text, x+12.5 y70 Section, Notes:
Gui Add, Edit, w215 h120 vnotes, %notes%
Gui, Add, Button, w200 h30 gMyButton, Get Quote Info  ; Create a button

; Add a Text control
; Gui, Add, Text, vMyText, Salesforce Checklist Content

; Add a button that toggles the visibility of the Text control
Gui, Add, Button, gToggleVisibility, Toggle Checklists

;----------- START CHECKLISTS ---------------
Gui Add,Tab3, vsalesforceChecklist, Salesforce Checklist|SAP Checklist - Main Page|SAP Checklist - Inside The Order|SAP - Finalizing The Order
Gui Tab, 1

; ----------- PRE SALESFORCE -----------------

Gui Add, GroupBox, Section h175 w250, Pre Salesforce
Gui Add, Checkbox, xp+10 yp+30 gsubmitChecklist vnameCheck, Check TENA name on PO
Gui Add, Checkbox, gsubmitChecklist vquoteNumberMatch, Quote number matches on PO && quote
Gui Add, Checkbox, gsubmitChecklist vpaymentTerms, Payment terms match on PO && quote
Gui Add, Checkbox, gsubmitChecklist vpriceMatch, Prices match on PO && quote
Gui Add, Checkbox, gsubmitChecklist vbothAddresses, Bill to && Ship to address on PO

; ----------- END PRE SALESFORCE -----------------

; ----------- SALESFORCE -----------------

Gui Add, GroupBox, Section ys h175 w325, Salesforce
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vpdfQuote, Save PDF of Quote (Document Output Tab)
Gui Add, Checkbox, gsubmitChecklist varrangeLines, Arrange quote lines (if needed) (Quote Details Tab)
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
Gui Add, Checkbox, xs+10 y+10 gsubmitChecklist vdpsAttached, DPS Reports
Gui Add, Checkbox, x+5 gsubmitChecklist vorderNoticeAttached, Order Notice
Gui Add, Text, xs+10 y+10, WIN Form?
Gui Add, Radio, x+26 gsubmitChecklist vwinYes, Yes
Gui Add, Radio, x+5 gsubmitChecklist vwinNa, N/A
Gui Add, Text, xs+10 y+10, Merge Report?
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
Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vfinalTotal, Check Net Value
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

Gui Show, w920 h530, Order Organizer ;SO# %soNumber%
Gui Submit, NoHide

submitChecklist:
Gui Submit, Nohide
return

ToggleVisibility:
    ; Toggle the visibility of the Edit control
    if (checkListVisible)
		{
			GuiControl, Hide, salesforceChecklist
			checkListVisible := false
			Gui Show, w920 h530, Order Organizer ;SO# %soNumber%
		}
		else
		{
			GuiControl, Show, salesforceChecklist
			checkListVisible := true
			Gui Show, w920 h760, Order Organizer ;SO# %soNumber%
		}
return

; CalculateTotals:
; Gui Submit, NoHide
; StringReplace, poValue, poValue, `,,, All
; StringReplace, tax, tax, `,,, All
; StringReplace, surcharge, surcharge, `,,, All
; StringReplace, freightCost, freightCost, `,,, All


if (poValue)
{
	totalCost := poValue
}

if (poValue) && (tax)
{
	totalCost := poValue + tax
}

if (poValue) && (freightCost)
{
	totalCost := poValue + freightCost
}

if (poValue) && (surcharge)
{
	totalCost := poValue + surcharge
}

if (poValue) && (tax) && (freightCost)
{
	totalCost := poValue + tax + freightCost
}

if (poValue) && (tax) && (surcharge)
{
	totalCost := poValue + tax + surcharge
}

if (poValue) && (freightCost) && (surcharge)
{
	totalCost := poValue + freightCost + surcharge
}

if (poValue) && (tax) && (freightCost) && (surcharge)
{
	totalCost := poValue + tax + freightCost + surcharge
}

totalCost := Round(totalCost, 2)
GuiControl,,totalCost, %totalCost%
Return

submitSales:
Gui Submit, NoHide
gosub, findSales
return

dongle:
Gui Submit, NoHide
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
Gui destroy
Return

ClearFields:
Gui, Submit, NoHide
GuiControl,, cpq,
GuiControl,, po,
GuiControl,, sot,
GuiControl,, customer,
GuiControl,, contact,
GuiControl,, address,
GuiControl,, soldTo,
GuiControl,, terms,
GuiControl,, system,
GuiControl,, salesPerson,
GuiControl,, salesManager,
GuiControl,, managerCode,
GuiControl,, salesDirector,
GuiControl,, directorCode,
; GuiControl,, software,
GuiControl,, serialNumber,
GuiControl,, crd,
GuiControl,, poDate,
GuiControl,, sapDate,
GuiControl,, poValue,
GuiControl,, tax,
GuiControl,, freightCost,
GuiControl,, surcharge,
GuiControl,, totalCost,
GuiControl,, endUser,
GuiControl,, phone,
GuiControl,, email,
GuiControl,, endUse,
GuiControl,, soNumber,
GuiControl,, notes,

; Clearing checkboxes
GuiControl,, nameCheck, 0
GuiControl,, quoteNumberMatch, 0
GuiControl,, paymentTerms, 0
GuiControl,, priceMatch, 0
GuiControl,, bothAddresses, 0
GuiControl,, pdfQuote, 0
GuiControl,, arrangeLines, 0
GuiControl,, soldToIdCheck, 0
GuiControl,, orderTypeCheck, 0
GuiControl,, poInfoCheck, 0
GuiControl,, generateDps, 0
GuiControl,, orderNoticeSent, 0
GuiControl,, enteredSot, 0
GuiControl,, reorderCheck, 0
GuiControl,, backOrderCheck, 0
GuiControl,, creditCardCheck, 0
GuiControl,, shipCompleteCheck, 0
GuiControl,, mergeLinesCheck, 0
GuiControl,, shipDateCheck, 0
GuiControl,, taxCheck, 0
GuiControl,, freightChargeCheck, 0
GuiControl,, miscellaneousChargeCheck, 0
GuiControl,, termsDiscountCheck, 0
GuiControl,, poAttached, 0
GuiControl,, quoteAttached, 0
GuiControl,, dpsAttached, 0
GuiControl,, orderNoticeAttached, 0
GuiControl,, crdDateAdded, 0
GuiControl,, firstDate, 0
GuiControl,, incoterms, 0
GuiControl,, volts, 0
GuiControl,, shipper, 0
GuiControl,, verifyIncoterms, 0
GuiControl,, managerCodeCheck, 0
GuiControl,, directorCodeCheck, 0
GuiControl,, billtoCheck, 0
GuiControl,, shiptoCheck, 0
GuiControl,, contactPersonCheck, 0
GuiControl,, textsContactCheck, 0
GuiControl,, finalTotal, 0

; Clearing radio buttons
GuiControl,, tandcYes, 0
GuiControl,, tandcNo, 0
GuiControl,, tandcNa, 0
GuiControl,, winYes, 0
GuiControl,, winNo, 0
GuiControl,, winNa, 0
GuiControl,, mergeYes, 0
GuiControl,, mergeNo, 0
GuiControl,, mergeNa, 0
GuiControl,, addressTypeBilling, 0
GuiControl,, addressTypeShipping, 0
GuiControl,, addressTypeBoth, 0
GuiControl,, orderTypeStandard, 0
GuiControl,, orderTypeRush, 0
GuiControl,, orderTypeBackOrder, 0
GuiControl,, endUserCheck, 0
GuiControl,, endUserNA, 0
GuiControl,, serialYes, 0
GuiControl,, serialNa, 0
GuiControl,, endUserCopyBack, 0
GuiControl,, endUserCopyBackNa, 0
GuiControl,, shippingYes, 0
GuiControl,, shippingNa, 0
GuiControl,, higherLevelLinkingNa, 0
GuiControl,, higherLevelLinkingYes, 0
GuiControl,, deliveryGroupsYes, 0
GuiControl,, deliveryGroupsNa, 0
GuiControl,, updateDeliveryBlock, 0
GuiControl,, orderAcceptedNa, 0
GuiControl,, orderAcceptedYes, 0
return

Search:
GuiControlGet, SearchTerm
LV_Delete() ; Clear the ListView

; If SearchTerm is empty, add all files to the ListView
; Otherwise, only add files that contain the SearchTerm
for Each, FileName in LVArray {
    if (SearchTerm = "" or InStr(FileName, SearchTerm)) {
        LV_Add("", FileName)
    }
}

; Hide the ListView if no SearchTerm is entered, otherwise show it and bring the GUI to the top
if (SearchTerm = "") {
    GuiControl, Hide, LV
} else {
    GuiControl, Show, LV
	; Gui, +AlwaysOnTop ; Add the AlwaysOnTop property when the ListView is shown
}
return

MyListView:
if (A_GuiEvent = "DoubleClick") {
    LV_GetText(RowText, A_EventInfo)  ; Get the text from the row's first field.
    SelectedFile := myinipath . "\" . RowText
    Send +{tab}{BackSpace}
    Gosub, readtheini
}
return

restartScript:
Gosub, SaveToIni
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
Gui Submit, NoHide
if (cpq) && (po)
	gosub, SaveToIni
; FileSelectFile, SelectedFile,r,%myinipath%, Open a file ;______ Needs to go back
if (ErrorLevel)
	{
		gosub, restartScript
		return
	}
IniRead, ID, %SelectedFile%, id, ID
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
GuiControl,, salesPerson, %salesPerson%
IniRead, salesManager, %SelectedFile%, orderInfo, salesManager
GuiControl,, salesManager, %salesManager%
IniRead, salesDirector, %SelectedFile%, orderInfo, salesDirector
GuiControl, Choose, salesDirector, %salesDirector%
IniRead, managerCode, %SelectedFile%, orderInfo, managerCode
GuiControl, ChooseString, managerCode, %managerCode%
IniRead, directorCode, %SelectedFile%, orderInfo, directorCode
GuiControl, ChooseString, directorCode, %directorCode%
IniRead, address, %SelectedFile%, orderInfo, address
GuiControl,, address, %address%
IniRead, terms, %SelectedFile%, orderInfo, terms
GuiControl,, terms, %terms%
IniRead, contact, %SelectedFile%, orderInfo, contact
GuiControl,, contact, %contact%
IniRead, poValue, %SelectedFile%, orderInfo, poValue
GuiControl,, poValue, %poValue%
IniRead, tax, %SelectedFile%, orderInfo, tax
GuiControl,, tax, %tax%
IniRead, freightCost, %SelectedFile%, orderInfo, freightCost
GuiControl,, freightCost, %freightCost%
IniRead, surcharge, %SelectedFile%, orderInfo, surcharge
GuiControl,, surcharge, %surcharge%	
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
if % notesDeescaped == "ERROR"
	notesDeescaped :=
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
WinSetTitle, Order Organizer,,Order Organizer - SO# %soNumber% `(DBID-%ID%`)
return

SaveToIni:
Gui Submit, NoHide
if (!cpq) || (!po)
{
	MsgBox, Please enter a quote and PO#.
	return
}
IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . ".ini"
IniFilePathWithSo := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . " SO# " . soNumber . ".ini"
IniFilePathWithNoCPQ := myinipath . "\PO " . po . " " . customer . " SO# " . soNumber . ".ini"

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

WriteIniVariables:
Gui submit, NoHide
IniWrite, %cpq%, %IniFilePath%, orderInfo, cpq
IniWrite, %po%, %IniFilePath%, orderInfo, po
IniWrite, %sot%, %IniFilePath%, orderInfo, sot
IniWrite, %customer%, %IniFilePath%, orderInfo, customer
IniWrite, %terms%, %IniFilePath%, orderInfo, terms
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
IniWrite, %surcharge%, %IniFilePath%, orderInfo, surcharge
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
	Loop, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB\*.*
		{
		LV_Add("", A_LoopFileName)
		LVArray.Push(A_LoopFileName)
		}
		LV_ModifyCol()
		GuiControl, hide, LV
; } else
; {
; 	MsgBox, there was an error
; 	Return
; }
return

CheckIfFolderExists:
if (RegExMatch(cpq, "(?:^00*)", quoteNumberCpq))
{
	folderPath = C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs\PO %po% %customer% - CPQ-%cpq%
} else if (RegExMatch(cpq, "(?:^[2].*)", quoteNumberSap))
{
	folderPath = C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs\PO %po% %customer% - Quote %cpq%
} else if (cpq != quoteNumberCpq || cpq != quoteNumbSap)
{
	MsgBox, Invalid Quote
}
if FileExist(folderPath)
	return
if !FileExist(folderPath)
	FileCreateDir, %folderPath%
	run, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs\
	; Return
return

SaveBar:
WinGetPos x, y, Width, Height, Order Organizer
	x += 350
	y += 50
myRange:=90
Gui 2: -Caption
Gui 2:+AlwaysOnTop
Gui 2: Color, default
Gui 2: Font, S8 cBlack, Segoe UI Semibold
Gui 2: Add, Text, x0 y1 w208 h16 +Center, Saving
Gui 2: Add,Progress, x10 y20 w190 h20 cblue vPro1 Range0-%myRange%,0
;~ Gui 1: Add, Button, x10 w150 h30 glooping, Loop over list
Gui 2: Show, w210 h50 x847 y313
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
Gui 2: Destroy
Return
}

orderInfo()


;-------- START TEXT SNIPPET MENU --------

; Create the popup menu by adding some items to it.
Menu, Snippets, Add, SAP SO#, SoNumber
Menu, Snippets, Add, CPQ or Quote#, Quote
Menu, Snippets, Add, PO, PO
Menu, Snippets, Add, Order Notice, OrderNotice
Menu, Snippets, Add, Customer Name, CustomerName
Menu, Snippets, Add, Customer Contact, ContactName
Menu, Snippets, Add, Customer Sold To Acct#, SoldTo
Menu, Snippets, Add, CRD, Crd
Menu, Snippets, Add  ; Add a separator line.
Menu, Snippets, Add, SalesPerson, SalesPerson

; Menu, Snippets, Add, Saleserson Code, SalesPersonCode
Menu, Snippets, Add, Sales Manager, SalesManager
Menu, Snippets, Add, Sales Manager Code, SalesManagerCode
Menu, Snippets, Add, Sales Director, SalesDirector
Menu, Snippets, Add, Sales Director Code, SalesDirectorCode

Menu, Snippets, Add  ; Add a separator line.
Menu, Snippets, Add, PO Value, PoValue
Menu, Snippets, Add, Freight Cost, FreightCost
Menu, Snippets, Add, Total Cost, TotalCost

; Menu, Snippets, Add, Surcharge, Surcharge

Menu, Snippets, Add  ; Add a separator line.
Menu, Snippets, Add, End User, EndUser
Menu, Snippets, Add, Phone#, Phone
Menu, Snippets, Add, Email, Email
Menu, Snippets, Add, End User Info, EndUserInfo

SoNumber:
Clipboard := soNumber
Send, ^v
Return

Quote:
Clipboard := cpq
Send, ^v
return

PO:
Clipboard := po
Send, ^v
return

CustomerName:
Clipboard := customer
Send, ^v
return

OrderNotice:
Clipboard := "Order Notice - " . customer . " - $" . poValue
Send, ^v
Return

SoldTo:
Clipboard := soldTo
Send, ^v
return

SalesPerson:
Clipboard := salesPerson
Send, ^v
return

SalesManager:
Clipboard := salesManager
Send, ^v
Return

SalesManagerCode:
Clipboard := managerCode
Send, ^v
return

; SalesPersonCode
; Clipboard := po
; Send, ^v
; return

SalesDirector:
Clipboard := salesDirector
Send, ^v
return

SalesDirectorCode:
Clipboard := directorCode
Send, ^v
return

ContactName:
Clipboard := contact
Send, ^v
return

PoValue:
Clipboard := poValue
Send, ^v
return

Crd:
FormatTime, TimeString, %crd%, MM/dd/yyyy
Clipboard := crd
Send, ^v
return

FreightCost:
Clipboard := freightCost
Send, ^v
return

TotalCost:
Clipboard := totalCost
Send, ^v
return

EndUser:
Clipboard := endUser
Send, ^v
return

Phone:
Clipboard := phone
Send, ^v
return

Email:
Clipboard := email
Send, ^v
return

EndUserInfo:
endUserInfo := "END USER: " . endUser . "`nPH: " . phone . "`nEMAIL: " . email . "`n`nCPQ-" . cpq . "`n`nEND USE: " . endUseDeescaped
StringUpper, endUserInfo, endUserInfo
Clipboard := endUserInfo
Send, ^v
return

#z::Menu, Snippets, Show  ; i.e. press the Win-Z hotkey to show the menu.

;******** HOTSTRINGS (TEXT EXPANSION) ********
#c::run calc.exe ; Run calculator
+PrintScreen::Send, +{F7} ; Next line in item Conditions SAP SOs
;----- Order keyboard shortcuts -----

; Alt::LWin
; LWin::Alt
; RWin::Alt



; RAlt::RWin
; RWin::RAlt

^Numpad7::
SendMode Event
Setkeydelay 80
Send ISURCHARGE{Tab}1{Tab}EA{Enter}
Sleep 1000
Gosub, WaitBeam
Send {Tab 5}{Up}{End}{BackSpace}0[{Enter}
Sleep 1000
Gosub, WaitBeam
Send {Up}^{Tab}{Tab 9}{Enter}
Gosub, WaitBeam
Sleep 1000
Send ^{End}
Sleep 500
Send PR00
SetKeyDelay 120
Send {Tab}^v{Home}{Delete}{Enter}
Gosub, WaitBeam
Send {F3}
Return
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
::zadd::
Send, %address%
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
:O:gsod::
Clipboard := "C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs"
Send ^v{Down}{enter}
Return
:O:pftl::
Send Hi ,`nPlease find the license attached for SO{#} %soNumber%.`n`nThank you^{home}{end}{left}
Return
::lv3::
MsgBox, 4, , Is this for Rodalyn?
IfMsgBox, Yes
{
    Clipboard := "Hi Rodalyn,`n`nPlease review for level 3 approval.`n`nThank you"
    Send % Clipboard
}
else
{
    Clipboard := "Hi Kathy,`n`nPlease review for level 3 approval.`n`nThank you"
    Send % Clipboard
}
; ::-x8::--------  --------
;----- End Order keyboard shortcuts -----

::emlinv::E-MAIL INVOICES TO:{space}
::sontc::Hi ,`nPlease see the SO notification below.`n`nThanks{up 4}{left}
; ::tenaa::THERMO ELECTRON NORTH AMERICA LLC
::zsig:: ; Default Email Signature
Send, !e2as{Enter}{down 10}{BackSpace 2}
return

::shrug::¯\_(ツ)_/¯ 
::winf::
Send, Hi %salesPerson%^+{left}{delete}{BackSpace},`nGreat order{!} Thank you for filling out the WIN form.`n`nRegards{up 3}{left}
return
::sohu::Hi ,`nThe SO has been updated.`n`nThanks{up 3}{left}
::orn::
Send, Order Notice - %customer% - $%poValue%
return


; /******** RMA HOTSTRINGS ********/
::sjr::PLEASE RETURN THESE ITEMS PREPAID TO:`nTHERMO FISHER SCIENTIFIC`nRMA {#}: XXXXXXX`n355 RIVER OAKS PARKWAY`nSAN JOSE, CA 95134{`n 2}PLEASE INCLUDE A COMPILED DECONTAMINATION FORM & A COPY OF THIS RETURN AUTHORIZATION
::rht::ORIGINAL SO{#} XXXXXX`nITEMS CONFIRMED DOWN BY FSE XXXXXXX{`n 2}REPLACEMENT SO{#} XXXXXX
::memr::PLEASE RETURN THESE ITEMS PREPAID TO:`nTHERMO FISHER SCIENTIFIC`nRMA {#}: XXXXXXX`n5025 Tuggle Rd`nMemphis, TN 38118-7514{`n 2}PLEASE INCLUDE A COMPILED DECONTAMINATION FORM & A COPY OF THIS RETURN AUTHORIZATION
;******** END HOTSTRINGS (TEXT EXPANSION) ********

ToAttachments:
	gosub,ToDisplaySap
	Send, ^+{Tab}!{down}
Return

if WinActive("Change Standard Order")
{
	PrintScreen::+F7
}


^\:: ;To main attachments Button in SAP
	Gosub, ToAttachments
return

#F3:: ; Search SOT by line#
	Send, ^g
	Sleep, 750
	Send, ^a{BackSpace}a%sot%{Enter}
return

!2:: ; Forward SW Licenses
	Send, Hi%firstname%`nPlease find the license attached for SO{#}{Space}%soNumber%.`n`nThanks{pgup}
return

^!v:: ; Show/Hide Order Info GUI
	DetectHiddenWindows, on
	; WinActivate, [ WinTitle, WinText, ExcludeTitle, ExcludeText]
	if !WinActive(Order Organizer, "Order Organizer")
	{
		WinActivate, Order Organizer, Order Organizer, ClipAngel
		return
	}
	else if WinActive(Order Organizer, "Order Organizer", ClipAngel)
	{
		WinMinimize
		return
	}
return

!#t:: ; Search Outlook Tasks
	Send, ^e
	KeyWait, enter, down
	Send, !js2f
return

^!s:: ;Save to SO Docs
	gosub, SaveToSoDocs
return

#5:: ; -------- Navigate to Exped Def 5 Day ---------
	Send, fah{left 12}
return

!':: ; Seach SAP By PO#
	Send, {down 6}
	Send, {Tab}{Enter}
return

!#u:: ; Update Order Status
	SetTitleMatchMode, 2
	Send, !ghs
	WinWait, Header Data, 
	IfWinNotActive, Header Data, , WinActivate, Header Data, 
		WinWaitActive, Header Data,
	Send, ^{tab 4}{enter}
	WinWait, Change Status, 
	IfWinNotActive, Change Status, , WinActivate, Change Status, 
		WinWaitActive, Change Status,
	Send, ^{tab 4}{Down}{Space}
return

!#h:: ;CSH Removal
	SendMode, Event
	SetKeyDelay, 50
	gosub WaitInbox
	gosub, GetSubjectFromOutlook
	ClipWait, 1
	RegExMatch(Clipboard, "(\d{7})", cshSoNumber)
	Sleep, 500
	Clipboard := cshSoNumber
	gosub, OpenSAPWindowForCsh
	WinWaitActive, Order %cshSoNumber%,,ClipAngel,
	Sleep, 100
	Send ^{tab 7}{down}{tab}{Delete}{Enter}
	gosub, WaitCshSO
	Sleep, 2000
	gosub, WaitInbox
	SetKeyDelay 0
	Send, ^+r
	gosub, GetSenderOrToFieldFromOutlook
	Send, Hi%firstname%,`n`nThe hold has been removed.`n`nThank you ; ^{Home}^{Right}^+{Right}+{F3 2}{Right}
return

; Save OA Checklist
!#k::
checklistPath := "C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Order Checklists\OA Checklist `- TEMPLATE.docx"
; Create an instance of Word
Word := ComObjCreate("Word.Application")

; Open the .docx file
Doc := Word.Documents.Open(checklistPath)
Word.Visible := True

; Save the active document with a specific file name and file format
myChecklistPaths := Word.ActiveDocument.SaveAs("C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Level 2 Approvals\OAC SO#" . soNumber . " - " . customer . ".docx")

; Get the range of the document
Range := Doc.Content

FormatTime, formattedDatePo, %poDate%, MM/dd/yyyy
FormatTime, formattedDateSap, %sapDate%, MM/dd/yyyy
FormatTime, formattedCrd, %crd%, MM/dd/yyyy

; Get a reference to the first table in the document
table := Word.ActiveDocument.Tables(1)

; Get a reference to the first cell in the table
cell := table.Cell(1, 2)
; Insert text into the cell
cell.Range.Text := customer

; Get a reference to the first cell in the table
cell := table.Cell(1, 6)
; Insert text into the cell
cell.Range.Text := soNumber

; Get a reference to the first cell in the table
cell := table.Cell(2, 2)
; Insert text into the cell
if !(address)
{
	cell.Range.Text := customer
}
else
{
	cell.Range.Text := address
}

; Get a reference to the first cell in the table
cell := table.Cell(2, 4)
; Insert text into the cell
cell.Range.Text := formattedDatePo

; Get a reference to the first cell in the table
cell := table.Cell(2, 6)
; Insert text into the cell
cell.Range.Text := formattedDateSap

range := Word.Selection.Range

; formattedCrd := "12/25/2023"
; paymentTerms := "Net 30"

; Get a reference to the 26th paragraph in the document
netTerms := Word.ActiveDocument.Range.Paragraphs(40).Range
; Set the range based on the found range
netTermsRange := netTerms
netTermsRange.Move(1,1)
; Set the text for the text box
netTermsRange.Text := terms

; Highlight the 26th paragraph
netTermsRange.Move(4,5)
netTermsRange.Move(1,1)
crdRange := netTermsRange
crdRange.Select
sleep 2000
crdRange.Text := formattedCrd

Doc.Close()

Word.Quit()

return
; SetTitleMatchMode, 3
; Run, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Order Checklists\OA Checklist - TEMPLATE.docx
; WinWaitActive, OA Checklist - TEMPLATE.docx - Word,
; Send, {f12}
; SetTitleMatchMode, 2
; WinWait, Save As, 
; IfWinNotActive, Save As, , WinActivate, Save As, 
; 	WinWaitActive, Save As,
; Sleep, 1000
; Send, +{tab 3}{Home}{down 2}
; Send, {Enter}
; Send, {tab}{space}
; Send, {Enter}
; Send, {tab 2}{End}^{Left}{Left}^+{Left}
; SendRaw, SO#
; Send, {Space}
return

^Numpad4::
	gosub, CopySOFromSubjectInOutlook
return

!#w:: ; Save from Word to SO Docs
	Send, {f12}
	WinWait, Save As, 
	IfWinNotActive, Save As, , WinActivate, Save As, 
		WinWaitActive, Save As,
	Sleep, 1000
	Send, +{tab 3}{Home}{down 2}
	Send, {Enter}
	Sleep, 500
	Send, {tab}{space}
return

WaitSaveAs:
	; WinWait, Save As, 
	; IfWinNotActive, Save As, , WinActivate, Save As, 
		WinWaitActive, Save As,
return

;~ JS / Chrome Test
;~ WinActivate, ahk_exe chrome.exe

OpenEmailAttachments:
	Emails := ComObjActive("Outlook.Application").ActiveExplorer.Selection
	for Email in Emails
		for Attachment in Email.Attachments
		if !(Attachment.FileName ~= "^image\d+") ; exlude image files often created by embedded images in email signatures
	{
		Attachment.SaveAsFile(A_Temp "\" Attachment.FileName)
		Run, % A_Temp "\" Attachment.FileName
	}
	WinWaitActive, % Attachment.Filename
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
Gui Show, x%X% y%Y% w300 h300
return

; ------------------------------------------

!+a:: ; Attach last file inside
	Send, !e2af
	Sleep 3000
	Send {Enter}
; return

!#a:: ; Attach last file pop out
	Send, !haf
	Sleep 3000
	Send {Enter}
; return

; ***** FUNCTIONS ***** | ***** FUNCTIONS ***** | ***** FUNCTIONS***** | ***** FUNCTIONS ***** 
; orderFocus() ; Display documents button on SAP Standard Order Page
; {
	; ControlFocus, Button1, Change Standard Order %soNumber%,
; }
; return
#!f::
; forwardSoftwareLicense()
Return

toMiddle() ; To the middle section of SAP Standard Order Page
{
	orderWindowActivate()
	Send, ^{Tab 3}
}
return

orderWindowActivate() ; Activate SAP Change Order Window
{
	WinActivate, Change Standard Order %soNumber%, 
}
return

; ***** END FUNCTIONS ***** | ***** END FUNCTIONS ***** | ***** END FUNCTIONS***** | ***** END FUNCTIONS ***** 

; ***** LABELS ***** | ***** LABELS ***** | ***** LABELS ***** | ***** LABELS ***** 


WaitSpin:
	Loop, ;Wait for mouse to not spin
	{
		if (A_Cursor = "AppStarting")
		{
			Loop, 
			{
				; MsgBox, In the loop
				if !(A_Cursor = "AppStarting")
					Break
			}
		}
	}
Return

WaitArrow:
	Loop, ;Wait for mouse to be arrow
	{
		Sleep, 500
		if (A_Cursor = "Arrow")
			Break
	}
return

WaitBeam:
	Loop, ;Wait for mouse to be arrow
	{
		Sleep, 750
		if (A_Cursor = "IBeam")|| (A_Cursor = "Arrow")
			Break
	}
return

OrderFocus:
	ControlFocus, Button1, Change Standard Order %soNumber%,
return

toMiddle:
	gosub, OrderWindowActivate
	Send, ^{Tab 3}
return

OrderWindowActivate:
	WinActivate, Change Standard Order %soNumber%, 
return

ToDisplaySap:
	ControlFocus, Button1, Standard Order
return

CopySOFromSubjectInOutlook:
	gosub, GetSubjectFromOutlook
	ClipWait, 1
	RegExMatch(Clipboard, "(\d{7})", soFromOutlook)
	Clipboard := soFromOutlook
return

ForwardEmail:
	WinActivate, Outlook
	olApp := ComObjCreate("Outlook.Application")
	olForward := olApp.ActiveExplorer.Selection.Item(1)
	olForward := olForward.Forward
	olForward.Display ; Remove to work in background
return

SaveToSoDocs: ; Go to SO Docs
	Send, {f12}
	WinWaitActive, Save As,
	Send, !d+{tab 10}{Home}{down 2}{ENTER}{Tab 3}
return

GetSubjectFromOutlook:
	; Get the subject of the active item in Outlook. Works in both the main window and
	; if the email is open in its own window.
	olApp := ComObjActive("Outlook.Application")
	olNamespace := olApp.GetNamespace("MAPI")
	; Get the active window so we can determine if an Explorer or Inspector window is active.
	Window := olApp.ActiveWindow 
	if (Window.Class = 34) { ; 34 = An Explorer object. (The Outlook main window)
		Selection := Window.Selection
		if (Selection.Count > 0)
			Clipboard := Selection.Item(1).Subject
		return
	}
	else
		return

GetSenderOrToFieldFromOutlook:
	Clipboard := % COMObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1).SenderName
	RegExMatch(Clipboard, "\s([a-zA-Z]*)", firstName)
return

; Find order by PO# in SAP
!#p::
	send, {tab 7}{enter}
return

WaitOrderSO:
	SetTitleMatchMode, 2
	WinActivate, Order %soNumber%,,ClipAngel,
	WinWaitActive, Order %soNumber%,,ClipAngel,
return

WaitCshSO:
	WinActivate, Order %cshSoNumber%,,ClipAngel,
	WinWaitActive, Order %cshSoNumber%,,ClipAngel,
Return

WaitHeaderData:
	SetTitleMatchMode, 2
	WinWait, Change Standard Order %soNumber%: Header Data,,ClipAngel,
	IfWinNotActive, Change Standard Order %soNumberso%: Header Data,,ClipAngel, WinActivate, Change Standard Order %soNumber%: Header Data, 
		WinWaitActive, Change Standard Order %soNumber%: Header Data,,ClipAngel,
return

WaitStandardOrder:
	SetTitleMatchMode, 2
	WinWait, Change Standard Order %soNumber%: Overview, 
	IfWinNotActive, Change Standard Order %soNumber%: Overview, , WinActivate,Change Standard Order %soNumber%: Overview, 
		WinWaitActive, Change Standard Order %soNumber%: Overview,
return

addAttachment: ; Add Attachment
	Send, !{down}
	Sleep, 1000
	Send, a
	Sleep, 2000
	Send, ^+{tab 3}{left 14}{enter}{down}{enter}
	Sleep, 500
	Send ^4
	Sleep, 500
	Send, !#s
return

WaitInbox:
	; WinWait, Inbox - matthew.terbeek@thermofisher.com - Outlook, 
	IfWinNotActive, Inbox - matthew.terbeek@thermofisher.com - Outlook, , WinActivate, Inbox - matthew.terbeek@thermofisher.com - Outlook, 
		WinWaitActive, Inbox - matthew.terbeek@thermofisher.com - Outlook, 
return

; ***** END LABELS ***** | ***** END LABELS ***** | ***** END LABELS ***** | ***** END LABELS ***** 

^1:: ; Create Attachment
	Gosub, OpenSAPWindow
	Gosub, ToAttachments
	Sleep, 500
	Send, c{Enter}
	Gosub, WinWaitImportFile
	Gosub, ToSoDocAttachments
	Send, !n
	Send, po{Space}%po%
	Sleep, 500
	Send, {Down}{Enter}
	Sleep, 1000
	Send, po
	Sleep, 500
	Send, {Down}{Enter}
	Sleep, 1000 
	Send, ^{tab 3}
	Sleep, 1500
	Send, !{Down}
	Sleep, 500
	Send, {down}a
	gosub, WinWaitAttachmentList
	Send, ^+{tab 3}{left 14}{enter}{down}{enter}
	WinWaitActive, Import file
	Send, !n
	gosub, WinWaitImportFile
	Send, cpq{Down}{enter}
	Sleep, 2000
	Send, ^+{tab}{left 14}{enter}{down}{enter}
	gosub, WinWaitImportFile
	Send, !n
	; Send, d
	; Sleep, 200
	; Send, {down}{enter}
	; Sleep, 2000
	; Send, ^+{tab}{left 14}{enter}
	; Sleep, 200
	; Send, {down}{enter}
	; gosub, WinWaitImportFile
	; Send, !n
	; Send, d
	; Sleep, 200
	; Send, {down 2}{enter}
Return

ToSoDocAttachments:
	WinWaitActive, Import file
	ControlFocus, ToolbarWindow322, Import file
	Send {space}
	Sleep 1000
	Send {tab}so{Enter}
	Sleep, 500
Return


WinWaitImportFile:
	WinWait, Import file, 
	WinActivate, Import file, 
	WinWaitActive, Import file,
return

WinWaitAttachmentList:
	WinWait, Service: Attachment list, 
	IfWinNotActive, Service: Attachment list, , WinActivate, Service: Attachment list, 
		WinWaitActive, Service: Attachment list, 
return

GetSOFromSubject:
	gosub, GetSubjectFromOutlook
	ClipWait, 1
	RegExMatch(Clipboard, "(\d{7})", so)
return

!#i:: ; Get Invoice - Not finished
WinWait, Change Billing Document, 
IfWinNotActive, Change Billing Document, , WinActivate, Change Billing Document, 
	WinWaitActive, Change Billing Document, 
Sleep, 100
Send, {ALTDOWN}{ALTUP}u
WinWait, Output output, 
IfWinNotActive, Output output, , WinActivate, Output output, 
	WinWaitActive, Output output, 
Send, {DOWN}{SHIFTDOWN}{SPACE}{SHIFTUP}{TAB 4}{ENTER}
WinWait, Print:, 
IfWinNotActive, Print:, , WinActivate, Print:, 
	WinWaitActive, Print:, 
Send, {CTRLDOWN}{SHIFTDOWN}{TAB}{SHIFTUP}{CTRLUP}{TAB}{ENTER}
WinWait, Print Preview, 
IfWinNotActive, Print Preview, , WinActivate, Print Preview, 
	WinWaitActive, Print Preview,
MouseClick, left, 1807, 170
WinWait,, Nitro Pro, 
IfWinNotActive,, Nitro Pro, WinActivate,,Nitro Pro,
	WinWaitActive,, Nitro Pro, 
Send, ^s
WinWait, Save As, 
IfWinNotActive, Save As, , WinActivate, Save As, 
		WinWaitActive, Save As, 
return

; OAC CHECKLIST

^l::
; Create an instance of Outlook
Outlook := ComObjCreate("Outlook.Application")

; Create a new mail item
Mail := Outlook.CreateItem(0)

; Ask the user to select a level
MsgBox, 4, Which Level?, Is this for Level 2?, 

IfMsgBox Yes
{
    ; Set the properties of the mail item for Level 2
    level := 2
	newEmail(Mail, level, soNumber)
}
Else
{
    ; Set the properties of the mail item for Level 3
	level := "2 and 3"
    newEmail(Mail, level, soNumber)
}

; Display the new mail item
Mail.Display()

return

newEmail(Mail, level, soNumber){
    ; Set the properties of the mail item
    Mail.Subject := "SO# " . soNumber . " Level " . level . " Approval"
	if level = 2
	{
		Mail.Body := "Hi Debbie,`n`nPlease review SO# " . soNumber .  " for level " . level . " approval.`n`nThank you"
	}
    else
	{
		Mail.Body := "Hi Debbie,`n`nPlease review SO# " . soNumber .  " and pass for level 3 approval.`n`nThank you"
	}
    Mail.To := "debbie.erickson@thermofisher.com"
}
Return

F15:: ; Copy / Paste - Plant Coding
	Sleep, 200
	Send, +{Home}+{Backspace}
	Send, ^v
	Send, {Down} 
	Send, {esc}{down}
return

+F15:: ; Delete coding
	Send, {End}+{Home}{backspace}{Esc}{Down}
return

!#s:: ;CPQ
	Send, !n
	sleep, 1000
	Send, cpq-{Down}{enter}
	Sleep, 2000
	Send, !d
	Sleep, 1000
return

#IfWinNotActive, Insert Hyperlink
{
	!x::
	Send +{F10},F,C ; Find related email in Outlook
	Return
}
return

^2:: ; Add'l Attachment
	Send, ^+{tab}{Home}{enter}{down}{enter}
	WinWait, Import file, 
	IfWinNotActive, Import file, , WinActivate, Import file, 
		WinWaitActive, Import file, 
	Send, !n
return

^+u:: ;---- To UPPERCASE ----
	Clipboard:= ""
	Sleep, 500
	Send, ^c ; copies selected text
	ClipWait, 1
	StringUpper Clipboard, Clipboard
	Send {Insert}
	Sleep 400
	Send ^v
	Sleep 400
	Send {Insert}
return

!+g:: ;---- Delete GSA Price --------
	Send, {tab}{tab}{del}{enter}
	Sleep, 500
	Send, +{F7}
return

!0:: ; Maximie all windows except Order Organizer & Sticky Notes
WinGet, MyCount, Count
GroupAdd, temp1,,,,Order Organizer
GroupAdd, AllWindows, ahk_group temp1,,,Sticky Notes
; GroupAdd, GroupName, WinTitle [, WinText, Label, ExcludeTitle, ExcludeText]
Loop, %MyCount%    {
	WinMaximize ahk_group AllWindows
  	Send !{tab}
}

OpenSAPWindow:
	SetTitleMatchMode, 2
	if WinExist("Change Standard Order") {
		WinActivate, Change Standard Order, Organizer ClipAngel
		WinWaitActive, Change Standard Order,, Organizer ClipAngel
		Send, {F3}{BackSpace}
		Sleep, 500
		Send, %soNumber%
		Sleep, 500
		Send, {Enter}
		Sleep, 500
		Send, {Tab}{Enter}
		WinActivate, Change Standard Order, Organizer ClipAngel
		WinWaitActive, Change Standard Order,, Organizer ClipAngel
	} else if WinExist("Change Sales Order") {
		WinActivate, Change Sales Order, Organizer ClipAngel
		WinWaitActive, Change Sales Order,, Organizer ClipAngel
		Sleep, 200
		Send, {end}+{Home}{BackSpace}%soNumber%{Enter}
		Sleep, 500
		Send, {Tab}{Enter}
		WinActivate, Change Standard Order, Organizer ClipAngel
		WinWaitActive, Change Standard Order,, Organizer ClipAngel
	} else if WinExist("SAP Easy Access") {
		IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access,
			WinWaitActive, SAP Easy Access,
		WinActivate
		MouseClick, left, 90, 66
		Send, va02{enter}
		WinWait, SAP Easy Access, 
		WinWaitActive, Change Sales Order,, Organizer ClipAngel
		Send, %soNumber%{enter}
		WinActivate, Change Standard Order, Organizer ClipAngel
		WinWaitActive, Change Standard Order,, Organizer ClipAngel
	}
return

OpenSAPWindowForCsh:
if WinExist("Change Standard Order") {
	WinActivate, Change Standard Order,
	WinWaitActive, Change Standard Order,
	Send, {F3}{BackSpace}
	Sleep, 500
	Send, ^v
	Sleep, 500
	Send, {Enter}
	Sleep, 500
	Send, {Tab}{Enter}
	IfWinNotActive, Change Standard Order, , WinActivate, Change Standard Order,
		WinWaitActive, Change Standard Order,
} else if WinExist("Change Sales Order") {
	WinActivate, Change Sales Order,
	WinWaitActive, Change Sales Order,
	Sleep, 200
	Send, {end}+{Home}{BackSpace}^v{Enter}
	Sleep, 500
	Send, {Tab}{Enter}
	IfWinNotActive, Change Standard Order, , WinActivate, Change Standard Order,
		WinWaitActive, Change Standard Order,
	WinActivate ; use the window found above
} else if WinExist("SAP Easy Access") {
	IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access,
		WinWaitActive, SAP Easy Access,
	WinActivate
	MouseClick, left, 90, 66
	Send, va02{enter}
	WinWait, SAP Easy Access, 
	IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access, 
		WinWaitActive, SAP Easy Access, 
	Send, ^v{enter}
	IfWinNotActive, Change Standard Order, , WinActivate, Change Standard Order,
		WinWaitActive, Change Standard Order,
	WinActivate ; use the window found above
}
Sleep 200
return

WaitSAPEasyAccess:
ifWinExist, SAP Easy Access
{
	IfWinNotActive, SAP Easy Access, , WinActivate, SAP Easy Access,
		WinWaitActive, SAP Easy Access,
	WinActivate
	return
} else
{
	MsgBox, Win Doesn't Exist
	return
}
return

!c:: ;DocuSign
Send, %customer%
KeyWait, F13, d
Send, {tab 2}{down}
Send, {tab 2}{down}{Tab}
Send, %manager%{tab}%salesPerson%{Tab}{Down}{Tab}%po%{Tab}%poValue%{Tab}{down}{Tab}CPQ-%cpq%{Tab}%soNumber%{Tab}c{Tab}n{tab}n{Tab}NET30{tab 3}
FormatTime, TimeString, %crd%, MM/dd/yyyy
Send, %TimeString%{tab}%contact% - %email%
return

!1:: ; CRD Date
FormatTime, TimeString, %crd%, MM/dd/yyyy
gosub, OrderFocus
gosub, ToMiddle
Send, {tab}%TimeString%{Enter}
Sleep, 1000
Loop, 10
{
	CoordMode, Pixel, Window
	ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\Screen_20211013092931.png
	If (ErrorLevel = 0)
	{
		Send, {Enter}
		WinWaitActive, Information
		WinWaitNotActive, Information
		break
	}
	Sleep, 500
}
WinWaitNotActive, Information
Sleep, 2000
Send, ^{tab}
Send, {tab 4}{enter}
Sleep, 2000
Send, !e
Send, st
WinWait, Change Delivery Date
; Loop, {
; 	IfWinActive, Change Delivery Date,
; 		ControlSend, Button2, {enter}, Change Delivery Date
; 	IfWinNotActive, Change Delivery Date, 
; 		break
; }
; Sleep, 2000
WinWaitNotActive, Standard Order: Availability Control
WinWaitActive, Change Standard Order %soNumber%: Overview,,ClipAngel,
Sleep, 2000
Send, +{F4} ; Go to Sales Header
WinWait, Change Standard Order %soNumber%: Header Data,,ClipAngel,
Send, ^{Tab}{Tab 2}{Down 2}{Right 2}{Enter} ; Sets 110 volts on sales tab
Sleep, 1000
Send, ^{pgdn}
return

!#m:: ;Merge Report
FormatTime, TimeString, %sapDate%, MM/dd/yyyy
sleep, 500
gosub, WaitSAPEasyAccess
Send, ^/
SendMode, event
SetKeyDelay, 70
Send, ZSOMRG{enter}
WinWait, IOLS Merge Drop Ship Report
Send, %soNumber%{tab 3}0020 ;{Down 9}
;~ Send, %TimeString%{enter}
;~ Send, {up 10}
;~ Send, +{end}^c
Send, {f8}
Sleep, 1000
Send, !l
Sleep, 750
Send, e
Sleep, 750
Send, a
Sleep, 750
Send, !n
Send, Merge Report - SO{#} %soNumber%
Send, {enter} 
return

#if WinActive("E-Mail Output Distribution List by Recipient")
{
	end::Send +{space}{down}
}
#if
return

^9:: ;Sales Employee 9
Send, ^+{end}
Sleep, 500
Send, +{tab}

if (directorCode = "N/A")
{
	Send, s{right 11}{Tab}
	Sleep 200
	Send, %managerCode%{enter}
}
else 
{
	Send, s{right 10}{Tab}
	Clipboard := directorCode
	Send % Clipboard
}
Sleep, 500
Send, ^{PGUP}
return

#Include QuoteInfo.ahk
MyButton:  ; Label for the button
	Gosub goGetQuoteInfo
	Gosub goGetWinForm
return

goGetQuoteInfo:
	getQuoteInfo(quoteID, contactName, contactEmail, contactPhone, customerName, quoteOwner, creatorManager, totalNetAmount, totalFreight, surcharge, totalTax, quoteTotal, soldToID, paymentTerms, opportunity)
	GuiControl,, cpq, %quoteID%
	GuiControl,, customer, %customerName%
	GuiControl,, contact, %contactName%
	GuiControl,, email, %contactEmail%
	GuiControl,, phone, %contactPhone%
	GuiControl,, address, %contactAddress%
	GuiControl,, soldTo, %soldToID%
	GuiControl,, salesPerson, %quoteOwner%
	GuiControl,, poValue, %totalNetAmount%
	GuiControl,, freightCost, %totalFreight%
	GuiControl,, surcharge, %surcharge%
	GuiControl,, tax, %totalTax%
	GuiControl,, totalCost, %quoteTotal%
	GuiControl,, salesPerson, %quoteOwner%
	GuiControl,, salesManager, %creatorManager%
	GuiControl,, terms, %paymentTerms%
	GuiControl,, system, %opportunity%
Return

goGetWinForm:
	getWinForm(opportunity, winFormLink, endUser, endUserPhoneNumber, endUserEmail, endUse)
	GuiControl,, endUser, %endUser%
	GuiControl,, phone, %endUserPhoneNumber%
	GuiControl,, email, %endUserEmail%
	GuiControl,, endUse, %endUse%
Return

^6:: ;End User Info
	endUserInfo = END USER: %endUser%`nPHONE: %phone%`nEMAIL: %email%`n`nEND USE: %endUse%`n`nCPQ-%cpq%
	StringUpper, endUserInfo, endUserInfo
	Clipboard := % endUserInfo
	gosub, WaitHeaderData
	Send, ^+{tab 3}{Right}{Enter} ; Navigate to texts tab
	Sleep, 1000
	MouseClick, left, 90,320
	Sleep, 1000
	Send, {PGDN}
	Sleep, 200
	Send, {Down 3} ; Navigate to End User Tab
	Sleep, 200 
	Send, ^{tab 2}{enter}
	Sleep, 500
	Send, %endUserInfo% ; Navigate to Entry textbox and paste
	Sleep, 200
	Send, {Up 4}{End}+{PGUP 3}^c
	ClipWait, 1
	Send, {F3}
	gosub, WaitStandardOrder
	Send, !g
	Sleep, 200
	Send, hp ; back to the partners tab
	gosub, WaitHeaderData
	Send, ^{end} ; back to end user
	Sleep, 500
	Send, {F2}
	Sleep, 200
	Send, {tab 3}{Enter}
	Sleep, 500
	Send, {tab 2}^v{Enter}
	Sleep, 500
	Send, ^{PGUP}
return

^!z::
	Send, ^s
	reload
return

kellerDropDown:	
	GuiControl, ChooseString, managerCode, 202375
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

mccormackDropDown:	
	GuiControl, ChooseString, managerCode, 1261
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

hewittDropDown:	
	GuiControl, ChooseString, managerCode, 98866
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

mcfaddenDropDown:	
	GuiControl, ChooseString, managerCode, 202610
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, N/A
	GuiControl, ChooseString, directorCode, N/A
	Gui Submit, NoHide
return

butlerDropDown:	
	GuiControl, ChooseString, managerCode, 1026
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

kleinDropDown:	
	GuiControl, ChooseString, managerCode, 1042
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

bennettDropDown:
	GuiControl, ChooseString, managerCode, 1416
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

rogersDropDown:
	GuiControl, ChooseString, managerCode, 203915
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

nadjieDropDown:
	GuiControl, ChooseString, managerCode, 96695
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

tollstrupDropDown:
	GuiControl, ChooseString, managerCode, 200320
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Sylveer Bergs
	GuiControl, ChooseString, directorCode, 203185
	Gui Submit, NoHide
return

craftsDropDown:
	GuiControl, ChooseString, managerCode, 202625
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

findSales:
;=============== Sales Managers ===================
if InStr(salesManager, "keller")
	gosub, kellerDropDown
if InStr(salesManager, "hewitt")
	gosub, hewittDropDown
if InStr(salesManager, "mccormack")
	gosub, mccormackDropDown
if InStr(salesManager, "butler")
	gosub, butlerDropDown
if InStr(salesManager, "klein")
	gosub, kleinDropDown
if InStr(salesManager, "nadjie")
	gosub, nadjieDropDown
if InStr(salesManager, "tollstrup")
	gosub, tollstrupDropDown
if InStr(salesManager, "crafts")
	gosub, craftsDropDown
if InStr(salesManager, "rogers")
	gosub, rogersDropDown
if InStr(salesManager, "bennett")
	gosub, bennettDropDown

if salesPerson =
{
	GuiControl, Choose, salesManager, |1 
	GuiControl, Choose, managerCode, |1
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, |1
	GuiControl, Choose, directorCode, |1
	Gui Submit, NoHide
}
return

!/::
Reload
Return

;======== KEYBOARD SHORTCUTS ========
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

;======== END KEYBOARD SHORTCUTS ========


NumpadClear::Pause
