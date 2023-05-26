#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
;~ #Include WatchFolder.ahk
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs ; Ensures a consistent starting directory.
#SingleInstance Force
#InstallKeybdHook
#KeyHistory 50
; #include <UIA_Interface>
; #include <UIA_Browser>

salesPeople := "|Justin Carder|Robin Sutka|Fred Simpson|Rhonda Oesterle|Mitch Lazaro|Tucker Lincoln|Jawad Pashmi|Julie Sawicki|Mike Hughes|"
. "Steve Boyanoski|Bob Riggs|Chuck Costanza|Navette Shirakawa|Stephanie Koczur|Mark Krigbaum|Jon Needels|Bill Balsanek|Brent Boyle|Andrew Clark"
. "|Kevin Clodfelter|Gabriel Mendez|Karl Kastner|Michael Burnett|Jerry Pappas|Nick Duczak|Steven Danielson|Nick Hubbard (Nik)|"
. "Samantha Stikeleather|Drew Smillie|Jeff Weller|Jerry Holycross|Theresa Borio|Dan Ciminelli|Cynthia (Cindy) Spittler|Gwyn Trojan"
. "|Joel Stradtner|Don Rathbauer|Hillary Tennant|Melissa Chandler|Douglas Sears|Rashila Patel|Brian Thompson|Larry Bellan|Donna Zwirner"
. "|Kristen Luttner|Helen Sun|May Chou|Haris Dzaferbegovic|Brian Dowe|Mark Woodworth|Susan Bird|Giovanni Pallante|Alicia Arias"
. "|Dominique Figueroa|Jonathan McNally|Murray Fryman|Yan Chen|Jie Qian|Joe Bernholz|David Kage|David Scott|Todd Stoner|John Bailey|"
. "Katianna Pihakari|Jonathan Ferguson|Aeron Avakian|Luke Marty|Alexander James|Timothy Johnson|Yuriy Dunayevskiy|Susan Gelman"
. "|Cari Randles|Shijun (Simon) Sheng|Sean Bennett|Nelson Huang|Lorraine Foglio|Gerald Koncar|Lauren Fischer|Brian Luckenbill"
. "|Amy Allgower|Brandon Markle|Crystal Flowers|Douglas McDowell|Dante Bencivengo|Dana Stradtner|Justin Chang|Kate Lincoln|"
. "Angelito Nepomuceno|Patrick Bohman|Kristin Roberts|John Venesky|Sarah Jackson|Daniel Quinn|Eric Norviel|Lisa Kasper|Karla Esparza|"
. "Loris Fossir|Russ Constantineau|David Kusel|Taylor Graham|Jerome Johemko|Christie Baldizar|Christina Guintu|Ross Milam|Sunny Chen|David Hill"
. "|Elaine Miller|Gary Scharrer|Valerie Bruner|Laura Howell|Cecilia Snyder|Tatiana Valle Melendez|Annie Cantelmo|Sitara Chauhan|Doug Meinhart"
. "|Tori Milioni|Bob Myers|Paulette Parker|Jane-Marie Kowalski|Ronsar Eid|Brian Ridley|Krystina Simms|"

salesManagers := "|Anjou Keller|Joe Hewitt|Zee Nadjie|Doug McCormack|Natalie Foels|Tonya Second|Lou Gavino|Christopher Crafts|Joe McFadden|John Butler|Richard Klein|Ray Chen|Randy Porch"

salesDirectors := "|Denise Schwartz|Joann Purkerson|Maroun El Khoury|Jimmy Yuk|N/A"

salesCodes := "|201020|202375|96715|1261|98866|96695|96654|202625|202006|1076|95410|202610|1026|1042|202756|202611|1041|N/A"

myinipath = C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB

I_Icon = C:\Users\%A_UserName%\OneDrive - Thermo Fisher Scientific\Desktop\Auto Hot Key Scripts\list_check_checklist_checkmark_icon_181579.ico
IfExist, %I_Icon%
	Menu, Tray, Icon, %I_Icon%

Menu, FileMenu, Add

SetTitleMatchMode, 2

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

Gui Add, Edit, x+20 y19 w125 h27.5 vSearchTerm gSearch, ; LV Search
Gui Add, ListView, grid r5 w400 y50 x250 vLV gMyListView, ORDERS:

; Gui Add, Edit, x+20 y19 w125 h27.5 vSearchTerm gSearch2, ; DDL Search
Loop, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB\*.*
{
LV_Add("", A_LoopFileName)
LVArray.Push(A_LoopFileName)
}
GuiControl, hide, LV


; Gui Add, DDL, w400 y50 x250 vLV gMyListView, ORDERS:
; Loop, % LVArray.MaxIndex()
; 	DDLString .= "|" LVArray[A_Index]
Gui Add, Button, xm+475 ym+10 w70 greadtheini, O&pen
Gui Add, Button, x+25 w70 gSaveToIni, &Save
Gui Add, Button, x+25 w150 grestartScript, &New PO or Reload
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
Gui Add, Text,, Customer / DPS Address:
Gui Add, Edit, vaddress, %address% 
Gui Add, Text,, Sold To Account:
Gui Add, Edit, vsoldTo, %soldTo%
Gui Add, Text,, Payment Terms
Gui Add, Edit, vterms, %terms% 
gui Add, GroupBox, x+12.5 y100 w1 h400 ; vertical line
Gui Add, Text, x+12.5 y72 Section, System:
Gui Add, Edit, vsystem, %system% 
gui Add, Text,, Salesperson:
Gui Add, DropDownList, +Sort vsalesPerson gsubmitSales, % salesPeople
Gui Add, Text,, Sales Manager:
Gui Add, DDL, Disabled vsalesManager, % salesManagers
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
Gui Add, Edit, vpoValue gCalculateTotals, %poValue%
Gui Add, Text,, Tax:
Gui Add, Edit, vtax gCalculateTotals, %tax%
Gui Add, Text,, Freight Cost:
Gui Add, Edit, vfreightCost gCalculateTotals, %freightCost% 
Gui Add, Text,, Surcharge:
Gui Add, Edit, w135 vsurcharge gCalculateTotals, %surcharge%
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

;----------- START CHECKLISTS ---------------

Gui Add,Tab3,, Salesforce Checklist|SAP Checklist - Main Page|SAP Checklist - Inside The Order|SAP - Finalizing The Order
Gui Tab, 1

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

Gui Show,w920, Order Organizer ;SO# %soNumber%
Gui Submit, NoHide

submitChecklist:
Gui Submit, Nohide
return

CalculateTotals:
Gui Submit, NoHide
StringReplace, poValue, poValue, `,,, All
StringReplace, tax, tax, `,,, All
StringReplace, surcharge, surcharge, `,,, All
StringReplace, freightCost, freightCost, `,,, All


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

Search:
; Loop, % LVArray.MaxIndex()
; 	DDLString .= "|" LVArray[A_Index]
GuiControlGet, SearchTerm
If SearchTerm =
{
	GuiControl, Hide, LV
}
	
; GuiControl, -Redraw, LV
LV_Delete()
For Each, FileName In LVArray
{
   If (SearchTerm != "")
   {
	   	GuiControl, Show, LV
		
		If InStr(FileName, SearchTerm) ; for overall matching
	   		LV_Add("", FileName)
		; DDLString .= "|" LVArray[A_Index]
   } Else
   {
	   LV_Add("", FileName)
   }
   		
}

Search2:
; Loop, Read, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Info DB\*.*
; {
; 	LVArray.Push(A_LoopFileName)
; }
; MsgBox, %LVArray%
; Gui Add, DropDownList, vLV h250 w400 y50 x250 gMyListView, %DDLString% 
Gui Submit, NoHide
Return

MyListView:
if (A_GuiEvent = "DoubleClick")
{
	LV_GetText(RowText, A_EventInfo)  ; Get the text from the row's first field.
	SelectedFile := myinipath . "\" . RowText
	Send +{tab}{BackSpace}
	Gosub, readtheini
}
Return

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
; if (ErrorLevel)
; 	{
; 		gosub, restartScript
; 		return
; 	}
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
!#p::Pause
Return
::-x8::--------  --------
;----- End Order keyboard shortcuts -----

::emlinv::E-MAIL INVOICES TO:{space}
::sontc::Hi ,`nPlease see the SO notification below.`n`nThanks{up 4}{left}
::tenaa::THERMO ELECTRON NORTH AMERICA LLC
::zsig:: ; Default Email Signature
Send, !e2as{Enter}{down 10}{BackSpace 2}
return

; :O:etid::everytime7Die{!} ; O at the beginning removes trailing space
; :O:igans::I'vegotanewshirt1996{!} ; O at the beginning removes trailing space
::mttf::matthew.terbeek@thermofisher.com

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

^#h:: ;CSH Removal
	SendMode, Event
	SetKeyDelay, 200
	gosub WaitInbox
	gosub, GetSubjectFromOutlook
	ClipWait, 1
	RegExMatch(Clipboard, "(\d{7})", cshSoNumber)
	Sleep, 500
	Clipboard := cshSoNumber
	gosub, OpenSAPWindowForCsh
	WinActivate, Order %cshSoNumber%,,ClipAngel,
	; Send, !ghd
	; WinWait, Change Standard Order %cshSoNumber%: Header Data,,ClipAngel,
	; Send, ^c
	; ClipWait, 1
	; cshPo := Clipboard
	; Send, {F3}
	; gosub, WaitCshSO
	; gosub WaitInbox
	; Send ^!s
	; gosub WaitSaveAs
	; Send, PO ^v
	; Sleep 500
	; Send, {Down}{Enter}
	; Clipboard := "CSH Removal Email"
	; Send ^v{Enter}
	; gosub WaitCshSO
	; Sleep 500
	; Send ^\
	; Sleep 1500
	; Send {Down 2}{Enter}
	; gosub WinWaitAttachmentList
	; Send, ^+{tab 3}{Home}{Enter}{Down}{Enter}
	; Gosub, WinWaitImportFile
	; Send !n
	; Clipboard := "C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs" 
	; Send ^v{enter}
	; Clipboard := cshPo
	; Send PO{space}^v{down}{enter}
	; Gosub, WinWaitImportFile
	; SetKeyDelay 150
	; Send csh{down}{enter}
	; Sleep 500
	; While (A_cursor = "AppStarting")
	; {
	; 	sleep 100
	; }
	; Sleep 500
	; WinWaitActive, Service: Attachment list, 
	; Sleep 1500
	; Send ^{Tab}{Enter}
	WinWaitActive, Order %cshSoNumber%,,ClipAngel,
	Sleep, 1000
	Send ^{tab 7}{down}{tab}{Delete}{Enter}
	gosub, WaitCshSO
	Sleep, 2000
	gosub, WaitInbox
	SetKeyDelay 0
	Send, ^+r
	gosub, GetSenderOrToFieldFromOutlook
	Send, Hi%firstname%,`nThe hold has been removed.`n`nThank you ; ^{Home}^{Right}^+{Right}+{F3 2}{Right}
return

; Save OA Checklist
!#k::
	SetTitleMatchMode, 3
	Run, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\Order Checklists\OA Checklist - TEMPLATE.docx
	WinWaitActive, OA Checklist - TEMPLATE.docx - Word,
	Send, {f12}
	SetTitleMatchMode, 2
	WinWait, Save As, 
	IfWinNotActive, Save As, , WinActivate, Save As, 
		WinWaitActive, Save As,
	Sleep, 1000
	Send, +{tab 3}{Home}{down 2}
	Send, {Enter}
	Send, {tab}{space}
	Send, {Enter}
	Send, {tab 2}{End}^{Left}{Left}^+{Left}
	SendRaw, SO#
	Send, {Space}
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
	Send, !e2af{Enter}
; return

!#a:: ; Attach last file pop out
	Send, !haf{Enter}
; return

; ***** FUNCTIONS ***** | ***** FUNCTIONS ***** | ***** FUNCTIONS***** | ***** FUNCTIONS ***** 
; orderFocus() ; Display documents button on SAP Standard Order Page
; {
	; ControlFocus, Button1, Change Standard Order %soNumber%,
; }
; return

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
!#f::
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
	Send, d
	Sleep, 200
	Send, {down}{enter}
	Sleep, 2000
	Send, ^+{tab}{left 14}{enter}
	Sleep, 200
	Send, {down}{enter}
	gosub, WinWaitImportFile
	Send, !n
	Send, d
	Sleep, 200
	Send, {down 2}{enter}
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

; ^l::
; SendMode, event
; Setkeydelay 20
; gosub WaitInbox
; Send ^n
; WinWait Untitled
; Clipboard := "talia.oberling@thermofisher.com"
; Send %Clipboard%
; Sleep 750
; Send {tab 3}
; Clipboard := "SO{#}{space}" . soNumber . "{space}Level 2 Approval"
; Send %Clipboard%{tab}
; Clipboard := "Hi Talia, `nPlease review SO{#}{space}" . soNumber .  "{space}for level 2 approval.`n`nThank you"
; Send %Clipboard%
; Sleep 500
; Send !h 
; Sleep 500
; Send af{enter}
; SetKeyDelay 0
; Return

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
Sleep 1000
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
	Send ^v
	Sleep 200
	Send {down}+{Tab}
	Sleep 200
	Send, s{right 11}{Tab}
	Sleep 200
	Clipboard := managerCode
	Send, ^v{enter}
}
Sleep, 500
Send, ^{PGUP}
return

#Include DPS.ahk

!+d:: ; |********** DPS REPORTS **********|
; getDPSReports(cpq,customer,address,contact)
Return

SetDefaultMouseSpeed, 10
dpsPath := "C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs\PO " . po . " " . customer . " - CPQ-" . cpq
; run, https://hub.thermofisher.com/ip
; WinWait, GTC: Homepage - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
; WinWaitActive, GTC: Homepage - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
; Sleep 200

; 	Loop, 
; 	{
; 		CoordMode, Pixel, Window
; 		ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, *3 C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\dps.png
; 		If (ErrorLevel = 0)
; 		{
; 			break
; 		}
; 	}

; ; MouseClick, WhichButton [, X, Y, ClickCount, Speed, D|U, R]
; MouseClick, left, 290, 190, 1, 0
; MouseClick, left, 537, 250, 1, 0
; WinWait, DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
; MouseClick, left, 1210, 377
; Send, TENA-CPQ-%cpq%
; MouseClick, left, 310, 494
; Send, %customer%
; MouseClick, left, 541, 527
; Send, %address%
; MouseClick left, 1771, 281
WinWait, DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
Gosub, DPSResults
Gosub, ReportGenerate
Gosub, PrintDPS

; Contact DPS Report
WinWaitActive, DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
Send, {tab 5}{BackSpace}+{Tab}{BackSpace}%contact%+{tab 7}{enter}
WinWaitActive, DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window

Gosub, DPSResults
Gosub, ReportGenerate
Gosub, PrintDPS
WinWaitActive, DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
Send, ^w
return

ReportGenerate:
generateReport := 0
Loop
{
	; ToolTip, ReportGenerate
	CoordMode, Pixel, Client
	ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\dateCreated.png
	If ErrorLevel = 0
	{
		MouseClick,, 781, 271
		; Override success screen
		;   Tab generate
		;   Click results
		Loop 
		{
			CoordMode, Pixel, Client
			ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\tab7.png
			if ErrorLevel = 0
			{
				; MsgBox Found 7
				Send {tab 7}{Enter}
				Break
			}
			CoordMode, Pixel, Client
			ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\tab8.png
			if ErrorLevel = 0
			{
				; MsgBox Found 8
				Send {tab 8}{Enter}
				Break
			}
			
		}
		Break
	}
}
; ToolTip
Return

PrintDPS:
ToolTip, PrintDPS
SetKeyDelay, 50
WinActivate, DTSSearchResults
WinWaitActive, DTSSearchResults
Sleep 500
Send ^p
Sleep, 750
Send {enter}
WinWaitActive, Save As
; Clipboard := dpsPath 
dpsPath := "C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs\PO " . po . " " . customer . " - CPQ-" . cpq . "\"
if FileExist(dpsPath . "DPS - " . customer . ".pdf")
{
	dpsContact := dpsPath . "DPS - " . contact
	Clipboard := dpsContact
} else 
{
	dpsCustomer := dpsPath . "DPS - " . customer
	Clipboard := dpsCustomer
}
Sleep, 200
WinWaitActive Save As
Send ^v
; While, (Clipboard)
; {
; 	Sleep 500
; }
; Sleep 3000
; Send !s
SendMode Event
SetKeyDelay 400
Send {tab 3}{enter}
WinWaitActive, DTSSearchResults
SetKeyDelay 0
Send ^w
ToolTip
return

DPSResults:
ToolTip, DPSResults
Loop
{   
	; No records found
	CoordMode, Pixel, Screen
	ImageSearch, FoundX, FoundY, 2214, 215, 2866, 380, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\noRecordsFound.png
    If ErrorLevel = 0
	{
		Sleep, 200
		Send, {tab 2}{enter}
		Break
	}

	; No records found / Clear
	CoordMode, Pixel, Screen
	ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\clear.png
    If ErrorLevel = 0
	{
		Sleep, 200
		Send, {tab 2}{enter}
		Break
	}
	
	; Blocked
	CoordMode, Pixel, Window
	ImageSearch, FoundX, FoundY, 959, 529, 1292, 617, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\blocked.png
	If ErrorLevel = 0
	{
		MsgBox, 4, Found Records,Records WERE Found`nContinue?
		If MsgBox No
		{
			Break
		}
		; Else
		{
			WinWaitActive, DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
			MouseClick, Left, 899, 257,1,, D ; reset tab position to middle of screen
			Send {tab 2}{enter}
			Sleep 200
			Send +{tab}{down 5}{enter}
			Sleep 200
			Send {tab 3}{Enter}
			WinWaitActive, DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
			; MsgBox Is Tab -> Enter Correct here?
			;Need search successfully overridden screenshot
			; Send {tab}{Enter}
			; Winwait here?
			; Need wait for overridden screenshot here?
			; Overridden
			CoordMode, Pixel, Client
			ImageSearch, FoundX, FoundY, 966, 487, 1241, 566, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\overridden.png
			If ErrorLevel = 0
			{
				Sleep, 200
				Send, {tab}{enter}
			}
			Break
		}
	}
	; Overridden
	CoordMode, Pixel, Client
	ImageSearch, FoundX, FoundY, 966, 487, 1241, 566, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\overridden.png
	If ErrorLevel = 0
	{
		Sleep, 200
		Send, {tab 2}{enter}
		Break
	}
}
ToolTip
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
	GuiControl, Choose, salesManager, Anjou Keller
	GuiControl, ChooseString, managerCode, 202375
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

mccormackDropDown:
	GuiControl, Choose, salesManager, Doug McCormack
	GuiControl, ChooseString, managerCode, 1261
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

hewittDropDown:
	GuiControl, Choose, salesManager, Joe Hewitt
	GuiControl, ChooseString, managerCode, 98866
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

foelsDropDown:
	GuiControl, Choose, salesManager, Natalie Foels
	GuiControl, ChooseString, managerCode, 96715
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

secondDropDown:
	GuiControl, Choose, salesManager, Tonya Second
	GuiControl, ChooseString, managerCode, 95410
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

mcfaddenDropDown:
	GuiControl, Choose, salesManager, Joe McFadden
	GuiControl, ChooseString, managerCode, 202610
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, N/A
	GuiControl, ChooseString, directorCode, N/A
	Gui Submit, NoHide
return

butlerDropDown:
	GuiControl, Choose, salesManager, John Butler
	GuiControl, ChooseString, managerCode, 1026
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

kleinDropDown:
	GuiControl, Choose, salesManager, Richard Klein
	GuiControl, ChooseString, managerCode, 1042
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

nadjieDropDown:
	GuiControl, Choose, salesManager, Zee Nadjie
	GuiControl, ChooseString, managerCode, 96695
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

chenDropDown:
	GuiControl, Choose, salesManager, Ray Chen
	GuiControl, ChooseString, managerCode, 202756
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Jimmy Yuk
	GuiControl, ChooseString, directorCode, 202611
	Gui Submit, NoHide
return

porchDropDown:
	GuiControl, Choose, salesManager, Randy Porch
	GuiControl, ChooseString, managerCode, 1041
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

tollstrupDropDown:
	GuiControl, Choose, salesManager, Darren Tollstrup
	GuiControl, ChooseString, managerCode, 200320
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Sylveer Bergs
	GuiControl, ChooseString, directorCode, 203185
	Gui Submit, NoHide
return

gavinoDropDown:
	GuiControl, Choose, salesManager, Lou Gavino
	GuiControl, ChooseString, managerCode, 96654
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 202375
	Gui Submit, NoHide
return

craftsDropDown:
	GuiControl, Choose, salesManager, Christopher Crafts
	GuiControl, ChooseString, managerCode, 202625
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
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
if salesPerson = Shijun (Simon) Sheng
	gosub, nadjieDropDown
if salesPerson = Brian Luckenbill
	gosub, nadjieDropDown
if salesPerson = Amy Allgower
	gosub, nadjieDropDown
if salesPerson = Michael Burnett
	gosub, nadjieDropDown
if salesPerson = Karla Esparza
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
if salesPerson = Cynthia (Cindy) Spittler
	gosub, mccormackDropDown
if salesPerson = Fred Simpson
	gosub, mccormackDropDown
if salesPerson = Gwyn Trojan
	gosub, mccormackDropDown
if salesPerson = Nick Hubbard (Nik)
	gosub, mccormackDropDown
if salesPerson = Kristen Luttner
	gosub, mccormackDropDown
if salesPerson = Patrick Bohman
	gosub, mccormackDropDown
if salesPerson = Eric Norviel
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
if salesPerson =  Dante Bencivengo 
	gosub, secondDropDown
if salesPerson =  Justin Chang 
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
if salesPerson = Dana Stradtner
	gosub, mcfaddenDropDown
if salesPerson = Sarah Jackson
	gosub, mcfaddenDropDown
if salesPerson = Loris Fossir
	gosub, mcfaddenDropDown
if salesPerson = Mitch Lazaro
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
if salesPerson = Daniel Quinn
	gosub, butlerDropDown
if salesPerson = Robin Sutka
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
if salesPerson = Kate Lincoln
	gosub, kleinDropDown
if salesPerson = Russ Constantineau
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
if salesPerson = Angelito Nepomuceno
	gosub, porchDropDown
if salesPerson = John Venesky
	gosub, porchDropDown
if salesPerson = Kristin Roberts
	gosub, porchDropDown
if salesPerson = David Kusel 
	gosub, porchDropDown
;=========== END PORCH ==========

;========== TOLLSTRUP ==========

if salesPerson = Taylor Graham  
	gosub, tollstrupDropDown
if salesPerson = Jerome Johemko  
	gosub, tollstrupDropDown
if salesPerson = Christie Baldizar 
	gosub, tollstrupDropDown

;========== END TOLLSTRUP ==========

;========== GAVINO ==========

if salesPerson = Christina Guintu  
	gosub, gavinoDropDown
if salesPerson = Ross Milam
	gosub, gavinoDropDown
if salesPerson = Sunny Chen  
	gosub, gavinoDropDown
if salesPerson = David Hill
	gosub, gavinoDropDown
if salesPerson = Elaine Miller
	gosub, gavinoDropDown
if salesPerson = Gary Scharrer
	gosub, gavinoDropDown
if salesPerson = Valerie Bruner
	gosub, gavinoDropDown
if salesPerson = Laura Howell
	gosub, gavinoDropDown
if salesPerson = Cecilia Snyder
	gosub, gavinoDropDown
if salesPerson = Tatiana Valle Melendez
	gosub, gavinoDropDown

;========== END GAVINO ==========

;======== CRAFTS =======

if salesPerson = Annie Cantelmo
	gosub, craftsDropDown
if salesPerson = Sitara Chauhan
	gosub, craftsDropDown
if salesPerson = Doug Meinhart
	gosub, craftsDropDown
if salesPerson = Tori Milioni
	gosub, craftsDropDown
if salesPerson = Bob Myers
	gosub, craftsDropDown
if salesPerson = Paulette Parker
	gosub, craftsDropDown
if salesPerson = Jane-Marie Kowalski
	gosub, craftsDropDown
if salesPerson = Ronsar Eid
	gosub, craftsDropDown
if salesPerson = Brian Ridley
	gosub, craftsDropDown 
if salesPerson = Krystina Simms
	gosub, craftsDropDown

;======== END CRAFTS =======


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
