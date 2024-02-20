#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%



;----------- START CHECKLISTS ---------------

; Gui Add,Tab3,, Salesforce Checklist|SAP Checklist - Main Page|SAP Checklist - Inside The Order|SAP - Finalizing The Order
; Gui Tab, 1

; ; ----------- PRE SALESFORCE -----------------

; Gui Add, GroupBox,Section h175 w250, Pre Salesforce
; Gui Add, Checkbox, xp+10 yp+30 gsubmitChecklist vnameCheck, Check TENA name on PO
; Gui Add, Checkbox, gsubmitChecklist vquoteNumberMatch, Quote number matches on PO && quote
; Gui Add, Checkbox, gsubmitChecklist vpaymentTerms, Payment terms match on PO && quote
; Gui Add, Checkbox, gsubmitChecklist vpriceMatch, Prices match on PO && quote
; Gui Add, Checkbox, gsubmitChecklist vbothAddresses, Bill to && Ship to address on PO

; ; ----------- END PRE SALESFORCE -----------------

; ; ----------- SALESFORCE -----------------

; Gui Add, GroupBox, Section ys h175 w325, Salesforce
; Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vpdfQuote, Save PDF of Quote (Document Output Tab)
; Gui Add, Checkbox, gsubmitChecklist varrangeLines, Arrange quote lines (if needed) (Quote Details Tab)
; Gui Add, Checkbox, gsubmitChecklist vsoldToIdCheck, Sold-To ID (Customer Details Tab)
; Gui Add, Checkbox, gsubmitChecklist vorderTypeCheck, Check order type (Order Tab)
; Gui Add, Checkbox, gsubmitChecklist vpoInfoCheck, Add PO# / Add PO Value / Upload PO (Order Tab)
; Gui Add, Checkbox, gsubmitChecklist vgenerateDps, Generate && Attach DPS Reports (Attachments Tab)

; ; ----------- END SALESFORCE -----------------

; ; ----------- PRE SAP -----------------

; Gui Add, GroupBox, Section ys h175 w275, Pre SAP
; Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vorderNoticeSent, Order Notice Sent
; Gui Add, Checkbox, gsubmitChecklist venteredSot, Entered In SOT
; Gui Add, Text, tcs y+10, T&&Cs?
; Gui Add, Radio, x+5 gsubmitChecklist vtandcYes, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vtandcNa, N/A

; ; ---------- END PRE-SAP ----------

; Gui, Tab, 2
; ; ----------- SAP -----------------

; Gui Add, GroupBox,Section x+15 ys+10 h130 w215, SAP Attachments:
; Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vpoAttached, PO
; Gui Add, Checkbox, x+55 gsubmitChecklist vquoteAttached, Quote
; Gui Add, Checkbox, xs+10 y+10 gsubmitChecklist vdpsAttached, DPS Reports
; Gui Add, Checkbox, x+5 gsubmitChecklist vorderNoticeAttached, Order Notice
; Gui Add, Text, xs+10 y+10, WIN Form?
; Gui Add, Radio, x+26 gsubmitChecklist vwinYes, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vwinNa, N/A
; Gui Add, Text, xs+10 y+10, Merge Report?
; Gui Add, Radio, x+10 gsubmitChecklist vmergeYes, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vmergeNa, N/A

; Gui Add, GroupBox, Section ys h100 w290, Main Sales Tab
; Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vcrdDateAdded, CRD Date (Req delv date)
; Gui Add, Checkbox, y+10 gsubmitChecklist vfirstDate, First Date Lines Updated (Edit -> Fast Change)
; Gui Add, Checkbox, gsubmitChecklist vincoterms, Incoterms

; Gui, Tab, 3

; Gui Add, GroupBox, Section h100 w200, Inside the Order
; Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vvolts, 110 Volts (Sales Tab)
; Gui Add, Checkbox, gsubmitChecklist vshipper, Shipper (Shipping Tab)
; Gui Add, Checkbox, gsubmitChecklist vverifyIncoterms, Verify Incoterms (Billing Tab)

; Gui Add, GroupBox, ys w250 h175 Section, Partners Tab
; Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vmanagerCodeCheck, Manager Code (Sales Manager)
; Gui Add, Checkbox, gsubmitChecklist vdirectorCodeCheck, Director Code (Sales Employee 9)
; Gui Add, Checkbox, gsubmitChecklist vbilltoCheck, Verify Billing Address
; Gui Add, Checkbox, gsubmitChecklist vshiptoCheck, Verify Shipping Address
; Gui Add, Text,, End User Added
; Gui Add, Radio, x+25 gsubmitChecklist vendUserCheck, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vendUserNA, N/A
; Gui Add, Checkbox, xs+10 y+10 gsubmitChecklist vcontactPersonCheck, Contact Person Added

; Gui Add, GroupBox, ys Section h175 w275, Texts Tab
; Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vtextsContactCheck, Contact Person or End User Info Added
; Gui Add, Text,, Serial/Dongle Number Added To Form Header?
; Gui Add, Radio, xs+25 y+10 gsubmitChecklist vserialYes, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vserialNa, N/A
; Gui Add, Text, xs+10 y+10, End User info added in the partners tab?
; Gui Add, Radio, xs+25 y+10 gsubmitChecklist vendUserCopyBack, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vendUserCopyBackNa, N/A

; Gui, Tab, 4

; Gui Add, GroupBox,Section h175 w235, Final Steps:
; Gui Add, Checkbox, xs+10 ys+25 gsubmitChecklist vfinalTotal, Check Net Value
; Gui Add, Text,, Add Shipping?
; Gui Add, Radio, x+50 gsubmitChecklist vshippingYes, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vshippingNa, N/A
; Gui Add, Text, xs+10 y+10, Higher Level Linking?
; Gui Add, Radio, x+15 gsubmitChecklist vhigherLevelLinkingYes, Yes
; Gui Add, Radio, x+5 vhigherLevelLinkingNa, N/A
; Gui Add, Text, xs+10 y+10, Delivery Groups?
; Gui Add, Radio, x+39 vdeliveryGroupsYes gsubmitChecklist, Yes
; Gui Add, Radio, x+5 vdeliveryGroupsNa gsubmitChecklist, N/A
; Gui Add, Checkbox, xs+10 y+10 vupdateDeliveryBlock gsubmitChecklist, Update Delivery Block
; Gui Add, Text, xs+10 y+10, Order Acceptance
; Gui Add, Radio, x+35 vorderAcceptedYes gsubmitChecklist, Yes
; Gui Add, Radio, x+5 vorderAcceptedNa gsubmitChecklist, N/A

; ; ----------- END SAP -----------------

; Gui Show,w920, Order Organizer ;SO# %soNumber%
; Gui Submit, NoHide

; submitChecklist:
; Gui Submit, Nohide
; return


; dongle:
; Gui Submit, NoHide
; if (software == 0)
; 	{
; 		GuiControl, Hide, serialNumber
; 		GuiControl, Hide, serialNumberText
; 		GuiControl, Move, software, y368
; 	}
; if (software == 1)
; 	{
; 		GuiControl, Move, software, y345
; 		GuiControl, Show, serialNumber
; 		GuiControl, Show, serialNumberText
; 	}
; return

; GuiClose:
; Gui destroy
; Return