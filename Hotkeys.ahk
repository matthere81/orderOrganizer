
; ---- HOTKEY & HOT STRING LIST ----

;-------- START TEXT SNIPPET MENU --------

; Create the popup menu by adding some items to it.
Menu, HotstringsContext, Add, SAP SO#, SoNumber
Menu, HotstringsContext, Add, CPQ or Quote#, Quote
Menu, HotstringsContext, Add, PO, PO
Menu, HotstringsContext, Add, Order Notice, OrderNotice
Menu, HotstringsContext, Add, Customer Name, CustomerName
Menu, HotstringsContext, Add, Customer Contact, ContactName
Menu, HotstringsContext, Add, Customer Sold To Acct#, SoldTo
Menu, HotstringsContext, Add, CRD, Crd
Menu, HotstringsContext, Add  ; Add a separator line.
Menu, HotstringsContext, Add, SalesPerson, SalesPerson

; Menu, HotstringsContext, Add, Saleserson Code, SalesPersonCode
Menu, HotstringsContext, Add, Sales Manager, SalesManager
Menu, HotstringsContext, Add, Sales Manager Code, SalesManagerCode
Menu, HotstringsContext, Add, Sales Director, SalesDirector
Menu, HotstringsContext, Add, Sales Director Code, SalesDirectorCode

Menu, HotstringsContext, Add  ; Add a separator line.
Menu, HotstringsContext, Add, PO Value, PoValue
Menu, HotstringsContext, Add, Freight Cost, FreightCost
Menu, HotstringsContext, Add, Total Cost, TotalCost

; Menu, HotstringsContext, Add, Surcharge, Surcharge
Menu, HotstringsContext, Add  ; Add a separator line.
Menu, HotstringsContext, Add, Serial Number, SerialNumber
Menu, HotstringsContext, Add, End User, EndUser
Menu, HotstringsContext, Add, Phone#, Phone
Menu, HotstringsContext, Add, Email, Email
Menu, HotstringsContext, Add, End User Info, EndUserInfo

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

SerialNumber:
Clipboard := serialNumber
Send, ^v
Return

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


Return




; ---- END HOTKEY & HOT STRING LIST ----






