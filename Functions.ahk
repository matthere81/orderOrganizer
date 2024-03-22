
Item1Handler:
    MsgBox, You selected Item1.

readtheini:
Gui, Submit, NoHide
; if (cpq) && (po)
;     gosub, SaveToIniNoGui
FileSelectFile, SelectedFile,r,%myinipath%, Open a file
; if (ErrorLevel)
;     {
;         gosub, restartScript
;         return
;     }

fields := ["cpq", "po", "sot", "customer", "salesPerson", "salesManager", "salesDirector", "managerCode", "directorCode", "address", "contact", "poValue", "tax", "freightCost", "surcharge", "totalCost", "system", "soldTo", "crd", "soNumber", "terms", "poDate", "sapDate", "endUser", "phone", "email"]

    for index, field in fields {
        IniRead, value, %SelectedFile%, orderInfo, %field%
        if (field == "sot" && value == "ERROR")
            value := 
        GuiControl,, %field%, %value%
    }
