; version := 1.2

; salesDirectors := "|Denise Schwartz|Joann Purkerson|Maroun El Khoury|Jimmy Yuk|Sylveer Bergs|N/A"
; managerCode := "|201020|202375|96715|1261|98866|96695|96654|202625|202006|1076|95410|202610|1026|1042|202756|202611|1041|200320|203185|1416|203915|N/A"

; I_Icon = C:\%A_UserName%\Thermo Fisher Scientificlist_check_checklist_checkmark_icon_181579.ico
; IfExist, %I_Icon%
; 	Menu, Tray, Icon, %I_Icon%


; ; Include Icon
; FileInstall, C:\Users\A_UserName\OneDrive - Thermo Fisher Scientific\Desktop\Auto Hot Key Scripts\list_check_checklist_checkmark_icon_181579.ico, A_WorkingDir, 1
; I_Icon = C:\Users\A_UserName\OneDrive - Thermo Fisher Scientific\Desktop\Auto Hot Key Scripts\list_check_checklist_checkmark_icon_181579.ico
; IfExist, %I_Icon%
; 	Menu, Tray, Icon, %I_Icon%

; SetTitleMatchMode, 2

; MsgBox here

Return

; ; Clearing checkboxes
; GuiControl,, nameCheck, 0
; GuiControl,, quoteNumberMatch, 0
; GuiControl,, paymentTerms, 0
; GuiControl,, priceMatch, 0
; GuiControl,, bothAddresses, 0
; GuiControl,, pdfQuote, 0
; GuiControl,, arrangeLines, 0
; GuiControl,, soldToIdCheck, 0
; GuiControl,, orderTypeCheck, 0
; GuiControl,, poInfoCheck, 0
; GuiControl,, generateDps, 0
; GuiControl,, orderNoticeSent, 0
; GuiControl,, enteredSot, 0
; GuiControl,, reorderCheck, 0
; GuiControl,, backOrderCheck, 0
; GuiControl,, creditCardCheck, 0
; GuiControl,, shipCompleteCheck, 0
; GuiControl,, mergeLinesCheck, 0
; GuiControl,, shipDateCheck, 0
; GuiControl,, taxCheck, 0
; GuiControl,, freightChargeCheck, 0
; GuiControl,, miscellaneousChargeCheck, 0
; GuiControl,, termsDiscountCheck, 0
; GuiControl,, poAttached, 0
; GuiControl,, quoteAttached, 0
; GuiControl,, dpsAttached, 0
; GuiControl,, orderNoticeAttached, 0
; GuiControl,, crdDateAdded, 0
; GuiControl,, firstDate, 0
; GuiControl,, incoterms, 0
; GuiControl,, volts, 0
; GuiControl,, shipper, 0
; GuiControl,, verifyIncoterms, 0
; GuiControl,, managerCodeCheck, 0
; GuiControl,, directorCodeCheck, 0
; GuiControl,, billtoCheck, 0
; GuiControl,, shiptoCheck, 0
; GuiControl,, contactPersonCheck, 0
; GuiControl,, textsContactCheck, 0
; GuiControl,, finalTotal, 0

; ; Clearing radio buttons
; GuiControl,, tandcYes, 0
; GuiControl,, tandcNo, 0
; GuiControl,, tandcNa, 0
; GuiControl,, winYes, 0
; GuiControl,, winNo, 0
; GuiControl,, winNa, 0
; GuiControl,, mergeYes, 0
; GuiControl,, mergeNo, 0
; GuiControl,, mergeNa, 0
; GuiControl,, addressTypeBilling, 0
; GuiControl,, addressTypeShipping, 0
; GuiControl,, addressTypeBoth, 0
; GuiControl,, orderTypeStandard, 0
; GuiControl,, orderTypeRush, 0
; GuiControl,, orderTypeBackOrder, 0
; GuiControl,, endUserCheck, 0
; GuiControl,, endUserNA, 0
; GuiControl,, serialYes, 0
; GuiControl,, serialNa, 0
; GuiControl,, endUserCopyBack, 0
; GuiControl,, endUserCopyBackNa, 0
; GuiControl,, shippingYes, 0
; GuiControl,, shippingNa, 0
; GuiControl,, higherLevelLinkingNa, 0
; GuiControl,, higherLevelLinkingYes, 0
; GuiControl,, deliveryGroupsYes, 0
; GuiControl,, deliveryGroupsNa, 0
; GuiControl,, updateDeliveryBlock, 0
; GuiControl,, orderAcceptedNa, 0
; GuiControl,, orderAcceptedYes, 0
; return

; restartScript:
; Gosub, SaveToIni
; if ((!cpq) || (!po))
; {
;     Reload
; } Else
; {
;     Gosub, SaveToIni
;     Reload
; } 
; return

; readtheini:
; Gui, Submit, NoHide
; if (cpq) && (po)
;     gosub, SaveToIniNoGui
; FileSelectFile, SelectedFile,r,%myinipath%, Open a file
; if (ErrorLevel)
;     {
;         gosub, restartScript
;         return
;     }


; IniRead, endUseEscaped, %SelectedFile%, orderInfo, endUse
; ; Get back newline separated list.
; StringReplace, endUseDeescaped, endUseEscaped, ``n, `n, All
; StringReplace, endUseDeescaped, endUseDeescaped, ``r, `r, All
; GuiControl,, endUse, %endUseDeescaped%

; IniRead, notesEscaped, %SelectedFile%, orderInfo, notes
; ; Get back newline separated list.
; StringReplace, notesDeescaped, notesEscaped, ``n, `n, All
; StringReplace, notesDeescaped, notesDeescaped, ``r, `r, All
; if % notes == "ERROR"
;     notes := 
; GuiControl,, notes, %notesDeescaped%

; ; GuiControl,, title, Order Organizer - SO# %soNumber%
; ; WinSetTitle, Order Organizer,,Order Organizer - SO# %soNumber%, Standard Order
; return

; SaveToIni:
; Gui, Submit, NoHide
; if (!cpq) || (!po)
; {
;     MsgBox, Please enter a quote and PO#.
;     return
; }
; IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . ".ini"
; IniFilePathWithSo := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . " SO# " . soNumber . ".ini"
; if FileExist(IniFilePath) && (soNumber)
; {
;     gosub, WriteIniVariables
;     FileMove, %IniFilePath%, %IniFilePathWithSo% , 1
;     IniFilePath = %IniFilePathWithSO% 
;     Gosub, SaveBar
;     gosub, CheckIfFolderExists
;     return  
; }
; else if FileExist(IniFilePathWithSo)
; {
;     IniFilePath = %IniFilePathWithSO% 
;     Gosub, SaveBar
;     gosub, WriteIniVariables
;     return
; }
; else if FileExist(IniFilePath) && (!soNumber)
; {
;     gosub, WriteIniVariables
;     Gosub, SaveBar
;     gosub, CheckIfFolderExists
;     return
; }
; else if !FileExist(IniFilePath) && !FileExist(IniFilePathWithSo)
; {
;     if(soNumber)
;     {
;         IniFilePath = %IniFilePathWithSo%
;     } else {
;         IniFilePath = %IniFilePath%
;     }
;     gosub, WriteIniVariables
;     Gosub, SaveBar
;     gosub, CheckIfFolderExists
;     return
; }
; return

; SaveToIniNoGui:
; IniFilePath := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . ".ini"
; IniFilePathWithSo := myinipath . "\PO " . po . " CPQ-" . cpq . " " . customer . " SO# " . soNumber . ".ini"
; if FileExist(IniFilePath) && (soNumber)
; {
;     gosub, WriteIniVariables
;     FileMove, %IniFilePath%, %IniFilePathWithSo% , 1
;     IniFilePath = %IniFilePathWithSO% 
;     gosub, CheckIfFolderExists
;     return  
; }
; else if FileExist(IniFilePathWithSo)
; {
;     IniFilePath = %IniFilePathWithSO% 
;     gosub, WriteIniVariables
;     return
; }
; else if FileExist(IniFilePath) && (!soNumber)
; {
;     gosub, WriteIniVariables
;     gosub, CheckIfFolderExists
;     return
; }
; else if !FileExist(IniFilePath) && !FileExist(IniFilePathWithSo)
; {
;     if(soNumber)
;     {
;         IniFilePath = %IniFilePathWithSo%
;     } else {
;         IniFilePath = %IniFilePath%
;     }
;     gosub, WriteIniVariables
;     gosub, CheckIfFolderExists
;     return
; }
; return

; WriteIniVariables:
; gui, submit, NoHide
; IniWrite, %cpq%, %IniFilePath%, orderInfo, cpq
; IniWrite, %po%, %IniFilePath%, orderInfo, po
; IniWrite, %sot%, %IniFilePath%, orderInfo, sot
; IniWrite, %customer%, %IniFilePath%, orderInfo, customer
; IniWrite, %salesPerson%, %IniFilePath%, orderInfo, salesPerson
; IniWrite, %salesManager%, %IniFilePath%, orderInfo, salesManager
; IniWrite, %managerCode%, %IniFilePath%, orderInfo, managerCode
; IniWrite, %salesDirector%, %IniFilePath%, orderInfo, salesDirector
; IniWrite, %directorCode%, %IniFilePath%, orderInfo, directorCode
; IniWrite, %address%, %IniFilePath%, orderInfo, address
; IniWrite, %contact%, %IniFilePath%, orderInfo, contact
; IniWrite, %poValue%, %IniFilePath%, orderInfo, poValue
; IniWrite, %tax%, %IniFilePath%, orderInfo, tax
; IniWrite, %freightCost%, %IniFilePath%, orderInfo, freightCost
; IniWrite, %surcharge%, %IniFilePath%, orderInfo, surcharge
; IniWrite, %totalCost%, %IniFilePath%, orderInfo, totalCost
; IniWrite, %system%, %IniFilePath%, orderInfo, system
; IniWrite, %soldTo%, %IniFilePath%, orderInfo, soldTo
; IniWrite, %crd%, %IniFilePath%, orderInfo, crd
; IniWrite, %soNumber%, %IniFilePath%, orderInfo, soNumber
; IniWrite, %terms%, %IniFilePath%, orderInfo, terms
; IniWrite, %poDate%, %IniFilePath%, orderInfo, poDate
; IniWrite, %sapDate%, %IniFilePath%, orderInfo, sapDate
; IniWrite, %endUser%, %IniFilePath%, orderInfo, endUser
; IniWrite, %phone%, %IniFilePath%, orderInfo, phone
; IniWrite, %email%, %IniFilePath%, orderInfo, email

;     ; Escape all newlines before writing it to ini file.
;     StringReplace, endUseEscaped, endUse, `n, ``n, All
;     StringReplace, endUseEscaped, endUseEscaped, `r, ``r, All
;     IniWrite, %endUseEscaped%, %IniFilePath%, orderInfo, endUse

;     ; Escape all newlines before writing it to ini file.
;     StringReplace, notesEscaped, notes, `n, ``n, All
;     StringReplace, notesEscaped, notesEscaped, `r, ``r, All
;     IniWrite, %notesEscaped%, %IniFilePath%, orderInfo, notes

; IniWrite, %softwareUpgradeYes%, %IniFilePath%, orderInfo, softwareUpgradeYes
; IniWrite, %softwareUpgradeNo%, %IniFilePath%, orderInfo, softwareUpgradeNo
; IniWrite, %softwareUpgradeLicense%, %IniFilePath%, orderInfo, softwareUpgradeLicense

; IniWrite, %software%, %IniFilePath%, orderInfo, software
; IniWrite, %serialNumber%, %IniFilePath%, orderInfo, serialNumber
; if (software == 0)
;     IniDelete, %IniFilePath%, orderInfo, serialNumber
; IniWrite, %nameCheck%, %IniFilePath%, orderInfo, nameCheck
; IniWrite, %orderNoticeSent%, %IniFilePath%, orderInfo, orderNoticeSent
; IniWrite, %enteredSot%, %IniFilePath%, orderInfo, enteredSot
; IniWrite, %tandcYes%, %IniFilePath%, orderInfo, tandcYes
; IniWrite, %tandcNa%, %IniFilePath%, orderInfo, tandcNa
; IniWrite, %poAttached%, %IniFilePath%, orderInfo, poAttached
; IniWrite, %quoteAttached%, %IniFilePath%, orderInfo, quoteAttached
; IniWrite, %dpsAttached%, %IniFilePath%, orderInfo, dpsAttached
; IniWrite, %orderNoticeAttached%, %IniFilePath%, orderInfo, orderNoticeAttached
; IniWrite, %winYes%, %IniFilePath%, orderInfo, winYes
; IniWrite, %winNa%, %IniFilePath%, orderInfo, winNa
; IniWrite, %mergeYes%, %IniFilePath%, orderInfo, mergeYes
; IniWrite, %mergeNa%, %IniFilePath%, orderInfo, mergeNa
; IniWrite, %finalTotal%, %IniFilePath%, orderInfo, finalTotal
; IniWrite, %shippingYes%, %IniFilePath%, orderInfo, shippingYes
; IniWrite, %shippingNa%, %IniFilePath%, orderInfo, shippingNa
; IniWrite, %higherLevelLinkingYes%, %IniFilePath%, orderInfo, higherLevelLinkingYes
; IniWrite, %higherLevelLinkingNa%, %IniFilePath%, orderInfo, higherLevelLinkingNa
; IniWrite, %deliveryGroupsYes%, %IniFilePath%, orderInfo, deliveryGroupsYes
; IniWrite, %deliveryGroupsNa%, %IniFilePath%, orderInfo, deliveryGroupsNa
; IniWrite, %updateDeliveryBlock%, %IniFilePath%, orderInfo, updateDeliveryBlock
; IniWrite, %orderAcceptedYes%, %IniFilePath%, orderInfo, orderAcceptedYes
; IniWrite, %orderAcceptedNa%, %IniFilePath%, orderInfo, orderAcceptedNa
; IniWrite, %serialYes%, %IniFilePath%, orderInfo, serialYes
; IniWrite, %serialNa%, %IniFilePath%, orderInfo, serialNa
; IniWrite, %endUserYes%, %IniFilePath%, orderInfo, endUserYes
; IniWrite, %endUserNa%, %IniFilePath%, orderInfo, endUserNa

; IniWrite, %quoteNumberMatch%, %IniFilePath%, orderInfo, quoteNumberMatch
; IniWrite, %paymentTerms%, %IniFilePath%, orderInfo,paymentTerms 
; IniWrite, %priceMatch%, %IniFilePath%, orderInfo, priceMatch
; IniWrite, %bothAddresses%, %IniFilePath%, orderInfo, bothAddresses

; IniWrite, %pdfQuote%, %IniFilePath%, orderInfo, pdfQuote
; IniWrite, %arrangeLines%, %IniFilePath%, orderInfo, arrangeLines
; IniWrite, %soldToIdCheck%, %IniFilePath%, orderInfo, soldToIdCheck
; IniWrite, %orderTypeCheck%, %IniFilePath%, orderInfo, orderTypeCheck
; IniWrite, %poInfoCheck%, %IniFilePath%, orderInfo, poInfoCheck
; IniWrite, %generateDps%, %IniFilePath%, orderInfo, generateDps

; IniWrite, %crdDateAdded%, %IniFilePath%, orderInfo,crdDateAdded 
; IniWrite, %firstDate%, %IniFilePath%, orderInfo, firstDate
; IniWrite, %incoterms%, %IniFilePath%, orderInfo, incoterms

; IniWrite, %volts%, %IniFilePath%, orderInfo, volts
; IniWrite, %shipper%, %IniFilePath%, orderInfo, shipper
; IniWrite, %verifyIncoterms%, %IniFilePath%, orderInfo, verifyIncoterms
; IniWrite, %managerCodeCheck%, %IniFilePath%, orderInfo, managerCodeCheck
; IniWrite, %directorCodeCheck%, %IniFilePath%, orderInfo, directorCodeCheck
; IniWrite, %billtoCheck%, %IniFilePath%, orderInfo, billtoCheck
; IniWrite, %shiptoCheck%, %IniFilePath%, orderInfo, shiptoCheck
; IniWrite, %endUserCheck%, %IniFilePath%, orderInfo, endUserCheck
; IniWrite, %endUserNA%, %IniFilePath%, orderInfo, endUserNA
; IniWrite, %contactPersonCheck%, %IniFilePath%, orderInfo, contactPersonCheck
; IniWrite, %textsContactCheck%, %IniFilePath%, orderInfo, textsContactCheck
; IniWrite, %endUserCopyBack%, %IniFilePath%, orderInfo, endUserCopyBack
; IniWrite, %endUserCopyBackNa%, %IniFilePath%, orderInfo, endUserCopyBackNa

; return

; CheckIfFolderExists:
; if (RegExMatch(cpq, "(?:^00*)", quoteNumberCpq))
; {
;     folderPath := A_WorkingDir . "\Order Docs\PO " . po . " " . customer . " - CPQ-" . cpq
; } else if (RegExMatch(cpq, "(?:^[2].*)", quoteNumberSap))
; {
;     folderPath := A_WorkingDir . "\Order Docs\PO %po% %customer% - Quote %cpq%"

; } else if (cpq != quoteNumberCpq || cpq != quoteNumbSap)
; {
;     MsgBox, Invalid Quote
; }
; if FileExist(folderPath)
;     return
; if !FileExist(folderPath)
;     FileCreateDir, %folderPath%
;     Return
; return

; SaveBar:
; WinGetPos x, y, Width, Height, Order Organizer
;     x += 350
;     y += 50
; myRange:=90
; Gui, 2: -Caption
; Gui, 2:+AlwaysOnTop
; Gui, 2: Color, default
; Gui, 2: Font, S8 cBlack, Segoe UI Semibold
; Gui, 2: Add, Text, x0 y1 w208 h16 +Center, Saving
; Gui, 2: Add,Progress, x10 y20 w190 h20 cblue vPro1 Range0-%myRange%,0
; ;~ Gui,1: Add, Button, x10 w150 h30 glooping, Loop over list
; Gui, 2: Show, w210 h50 x847 y313
; Gosub, Looping
; Return

; Looping:
; List1:=[1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0]
; Pro1:=0
; GuiControl,2:,Pro1,% Pro1
; Temp:=""
; Loop, 20
;     {
;         Temp.=List1[A_Index]
;         Pro1 +=10
;         GuiControl,2:,Pro1,% Pro1
;         sleep, 20
;     }
; Gui, 2: Destroy
; Return
; }

; orderInfo()

; ;-------- START TEXT SNIPPET MENU --------

; ; Create the popup menu by adding some items to it.
; Menu, Snippets, Add, SAP SO#, SoNumber
; Menu, Snippets, Add, CPQ or Quote#, Quote
; Menu, Snippets, Add, PO, PO
; Menu, Snippets, Add, Order Notice, OrderNotice
; Menu, Snippets, Add, Customer Name, CustomerName
; Menu, Snippets, Add, Customer Contact, ContactName
; Menu, Snippets, Add, Customer Sold To Acct#, SoldTo
; Menu, Snippets, Add, CRD, Crd
; Menu, Snippets, Add  ; Add a separator line.
; Menu, Snippets, Add, SalesPerson, SalesPerson

; ; Menu, Snippets, Add, Saleserson Code, SalesPersonCode
; Menu, Snippets, Add, Sales Manager, SalesManager
; Menu, Snippets, Add, Sales Manager Code, SalesManagerCode
; Menu, Snippets, Add, Sales Director, SalesDirector
; Menu, Snippets, Add, Sales Director Code, SalesDirectorCode

; Menu, Snippets, Add  ; Add a separator line.
; Menu, Snippets, Add, PO Value, PoValue
; Menu, Snippets, Add, Freight Cost, FreightCost
; Menu, Snippets, Add, Total Cost, TotalCost

; ; Menu, Snippets, Add, Surcharge, Surcharge

; Menu, Snippets, Add  ; Add a separator line.
; Menu, Snippets, Add, Serial Number, SerialNumber
; Menu, Snippets, Add, End User, EndUser
; Menu, Snippets, Add, Phone#, Phone
; Menu, Snippets, Add, Email, Email
; Menu, Snippets, Add, End User Info, EndUserInfo

; Return


; #Include QuoteInfo.ahk
; MyButton:  ; Label for the button
; 	Gosub goGetQuoteInfo
; 	Gosub goGetWinForm
; return

; goGetQuoteInfo:
; 	getQuoteInfo(quoteID, contactName, contactEmail, contactPhone, customerName, quoteOwner, creatorManager, totalNetAmount, totalFreight, surcharge, totalTax, quoteTotal, soldToID, paymentTerms, opportunity)
; 	GuiControl,, cpq, %quoteID%
; 	GuiControl,, customer, %customerName%
; 	GuiControl,, contact, %contactName%
; 	GuiControl,, email, %contactEmail%
; 	GuiControl,, phone, %contactPhone%
; 	; GuiControl,, address, %contactAddress%
; 	GuiControl,, soldTo, %soldToID%
; 	GuiControl,, salesPerson, %quoteOwner%
; 	GuiControl,, poValue, %totalNetAmount%
; 	GuiControl,, freightCost, %totalFreight%
; 	GuiControl,, surcharge, %surcharge%
; 	GuiControl,, tax, %totalTax%
; 	GuiControl,, totalCost, %quoteTotal%
; 	GuiControl,, salesPerson, %quoteOwner%
; 	GuiControl,, salesManager, %creatorManager%
; 	GuiControl,, terms, %paymentTerms%
; 	GuiControl,, system, %opportunity%
; Return

; goGetWinForm:
; 	getWinForm(opportunity, winFormLink, endUser, endUserPhoneNumber, endUserEmail, endUse)
; 	GuiControl,, endUser, %endUser%
; 	GuiControl,, phone, %endUserPhoneNumber%
; 	GuiControl,, email, %endUserEmail%
; 	GuiControl,, endUse, %endUse%
; Return



; SoNumber:
; Clipboard := soNumber
; Send, ^v
; Return

; Quote:
; Clipboard := cpq
; Send, ^v
; return

; PO:
; Clipboard := po
; Send, ^v
; return

; CustomerName:
; Clipboard := customer
; Send, ^v
; return

; OrderNotice:
; Clipboard := "Order Notice - " . customer . " - $" . poValue
; Send, ^v
; Return

; SoldTo:
; Clipboard := soldTo
; Send, ^v
; return

; SalesPerson:
; Clipboard := salesPerson
; Send, ^v
; return

; SalesManager:
; Clipboard := salesManager
; Send, ^v
; Return

; SalesManagerCode:
; Clipboard := managerCode
; Send, ^v
; return

; ; SalesPersonCode
; ; Clipboard := po
; ; Send, ^v
; ; return

; SalesDirector:
; Clipboard := salesDirector
; Send, ^v
; return

; SalesDirectorCode:
; Clipboard := directorCode
; Send, ^v
; return

; ContactName:
; Clipboard := contact
; Send, ^v
; return

; SerialNumber:
; Clipboard := serialNumber
; Send, ^v
; Return

; PoValue:
; Clipboard := poValue
; Send, ^v
; return

; Crd:
; FormatTime, TimeString, %crd%, MM/dd/yyyy
; Clipboard := crd
; Send, ^v
; return

; FreightCost:
; Clipboard := freightCost
; Send, ^v
; return

; TotalCost:
; Clipboard := totalCost
; Send, ^v
; return

; EndUser:
; Clipboard := endUser
; Send, ^v
; return

; Phone:
; Clipboard := phone
; Send, ^v
; return

; Email:
; Clipboard := email
; Send, ^v
; return

; EndUserInfo:
; endUserInfo := "END USER: " . endUser . "`nPH: " . phone . "`nEMAIL: " . email . "`n`nCPQ-" . cpq . "`n`nEND USE: " . endUseDeescaped
; StringUpper, endUserInfo, endUserInfo
; Clipboard := endUserInfo
; Send, ^v
; return

; #z::Menu, Snippets, Show  ; i.e. press the Win-Z hotkey to show the menu.

; ;******** HOTSTRINGS (TEXT EXPANSION) ********
; #c::run calc.exe ; Run calculator
; F13::Send, +{F7} ; Next line in item Conditions SAP SOs
; ;----- Order keyboard shortcuts -----
; ::zpo::
; Send, %po%
; return
; ::zpoz::
; Send, PO %po%
; return
; ::zso::
; Send, %soNumber%
; Return
; ::zsoz::
; Send, SO{#}{Space}%soNumber%
; return
; ::zpq::
; Send, %cpq%
; return
; ::zpqz::
; Send, CPQ-%cpq%
; return
; ::zval::
; Send, %poValue%
; return
; ::zsal::
; Send, %salesPerson%
; return
; ::zcust::
; Send, %customer%
; return
; ::zcon::
; Send, %contact%
; return
; ::zem::
; Send, %email%
; return
; ::zsys::
; Send, %system%
; return
; ::zenu::
; Send, %endUser%
; return
; ::zph::
; Send, %phone%
; return
; ::zuse::
; Send, %endUse%
; return
; ::zsot::
; Send, %sot% ;^{Left}{BackSpace}
; return
; :O:zcod::Close Out Document
; ::zcem::Contracts Email - 
; ::zwin::
; Send WIN Form - CPQ-%cpq%
; return
; ::ejim::10246281
; ;----- End Order keyboard shortcuts -----


; ;******** END HOTSTRINGS (TEXT EXPANSION) ********

; ^!v:: ; Show/Hide Order Info GUI
;     DetectHiddenWindows, on
;     if !WinActive(Order Organizer "ahk_class AutoHotkeyGUI")
;     {
;         WinActivate, Order Organizer ahk_class AutoHotkeyGUI, 
;         ;~ WinWaitActive, Order Info ahk_class AutoHotkeyGUI,
;         return
;     }
;     else if WinActive(Order Organizer "ahk_class AutoHotkeyGUI")
;     {
;         WinMinimize
;         return
;     }
; return

; ;/******** MONTIOR COUNT ********/

; SysGet, MonitorCount, MonitorCount
; SysGet, MonitorPrimary, MonitorPrimary
; MsgBox, Monitor Count:`t%MonitorCount%`nPrimary Monitor:`t%MonitorPrimary%
; Loop, %MonitorCount%
; {
;     SysGet, MonitorName, MonitorName, %A_Index%
;     SysGet, Monitor, Monitor, %A_Index%
;     SysGet, MonitorWorkArea, MonitorWorkArea, %A_Index%
;     MsgBox, Monitor:`t#%A_Index%`nName:`t%MonitorName%`nLeft:`t%MonitorLeft% (%MonitorWorkAreaLeft% work)`nTop:`t%MonitorTop% (%MonitorWorkAreaTop% work)`nRight:`t%MonitorRight% (%MonitorWorkAreaRight% work)`nBottom:`t%MonitorBottom% (%MonitorWorkAreaBottom% work)
; }

; X := 250, Y := 250 ; Starting position for the Gui on your main monitor
; CoordMode, Mouse, Screen
; MouseGetPos, MX, MY
; If (MX > A_ScreenWidth)
;     X += A_ScreenWidth
; Gui, Show, x%X% y%Y% w300 h300
; return

; kellerDropDown:
; 	GuiControl, ChooseString, managerCode, 202375
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Denise Schwartz
; 	GuiControl, ChooseString, directorCode, 201020
; 	Gui Submit, NoHide
; return

; mccormackDropDown:	
; 	GuiControl, ChooseString, managerCode, 1261
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Maroun El Khoury
; 	GuiControl, ChooseString, directorCode, 1076
; 	Gui Submit, NoHide
; return

; hewittDropDown:	
; 	GuiControl, ChooseString, managerCode, 98866
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Maroun El Khoury
; 	GuiControl, ChooseString, directorCode, 1076
; 	Gui Submit, NoHide
; return

; mcfaddenDropDown:	
; 	GuiControl, ChooseString, managerCode, 202610
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, N/A
; 	GuiControl, ChooseString, directorCode, N/A
; 	Gui Submit, NoHide
; return

; butlerDropDown:	
; 	GuiControl, ChooseString, managerCode, 1026
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Denise Schwartz
; 	GuiControl, ChooseString, directorCode, 201020
; 	Gui Submit, NoHide
; return

; kleinDropDown:	
; 	GuiControl, ChooseString, managerCode, 1042
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Maroun El Khoury
; 	GuiControl, ChooseString, directorCode, 1076
; 	Gui Submit, NoHide
; return

; bennettDropDown:
; 	GuiControl, ChooseString, managerCode, 1416
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Maroun El Khoury
; 	GuiControl, ChooseString, directorCode, 1076
; 	Gui Submit, NoHide
; return

; rogersDropDown:
; 	GuiControl, ChooseString, managerCode, 203915
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Denise Schwartz
; 	GuiControl, ChooseString, directorCode, 201020
; 	Gui Submit, NoHide
; return

; nadjieDropDown:
; 	GuiControl, ChooseString, managerCode, 96695
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Denise Schwartz
; 	GuiControl, ChooseString, directorCode, 201020
; 	Gui Submit, NoHide
; return

; tollstrupDropDown:
; 	GuiControl, ChooseString, managerCode, 200320
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Sylveer Bergs
; 	GuiControl, ChooseString, directorCode, 203185
; 	Gui Submit, NoHide
; return

; craftsDropDown:
; 	GuiControl, ChooseString, managerCode, 202625
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, Denise Schwartz
; 	GuiControl, ChooseString, directorCode, 201020
; 	Gui Submit, NoHide
; return

; findSales:
; ;=============== Sales Managers ===================
; if InStr(salesManager, "keller")
; 	gosub, kellerDropDown
; if InStr(salesManager, "hewitt")
; 	gosub, hewittDropDown
; if InStr(salesManager, "mccormack")
; 	gosub, mccormackDropDown
; if InStr(salesManager, "butler")
; 	gosub, butlerDropDown
; if InStr(salesManager, "klein")
; 	gosub, kleinDropDown
; if InStr(salesManager, "nadjie")
; 	gosub, nadjieDropDown
; if InStr(salesManager, "tollstrup")
; 	gosub, tollstrupDropDown
; if InStr(salesManager, "crafts")
; 	gosub, craftsDropDown
; if InStr(salesManager, "rogers")
; 	gosub, rogersDropDown
; if InStr(salesManager, "bennett")
; 	gosub, bennettDropDown

; if salesPerson =
; {
; 	GuiControl, Choose, salesManager, |1 
; 	GuiControl, Choose, managerCode, |1
; 	Gui Submit, NoHide
; 	GuiControl, Choose, salesDirector, |1
; 	GuiControl, Choose, directorCode, |1
; 	Gui Submit, NoHide
; }
; return