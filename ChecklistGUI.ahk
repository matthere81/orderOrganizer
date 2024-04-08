

;----------- START CHECKLISTS ---------------

Gui Add,Tab3,, Salesforce Checklist|SAP Checklist - Main Page|SAP Checklist - Inside The Order|SAP - Finalizing The Order
Gui Tab, 1

; ; Define the sections and their checklists
sections := ["Pre Salesforce", "Salesforce", "Pre SAP"]
checklists := [["Check TENA name on PO", "Quote number matches on PO && quote", "Payment terms match on PO && quote", "Prices match on PO && quote"
                , "Bill to && Ship to address on PO"]] ;,["Save PDF of Quote (Document Output Tab)", "Arrange quote lines (if needed) (Quote Details Tab)"
                ; , "Sold-To ID (Customer Details Tab)", "Check order type (Order Tab)", "Add PO# / Add PO Value / Upload PO (Order Tab)"]]
                ; , ["Order Notice Sent", "Entered In SOT", "T&&Cs?"]]

; Loop over each section
Loop, % sections.MaxIndex()
    {
        ; Add a group box for the section
        Gui Add, GroupBox, Section h175 w250, % sections[A_Index]

        ; Initialize the inner loop index
        innerIndex := 1
    
        ; Loop over each item in the checklist
        Loop, % checklists[A_Index].MaxIndex()
        {
            ; Get the checklist item
            item := checklists[A_Index][A_Index[]]
            MsgBox % item
            ; Create a sanitized version of the item to use as the variable name
            varName := StrReplace(item, " ", "_")
            varName := StrReplace(varName, "&", "_")
            varName := StrReplace(varName, "(", "_")
            varName := StrReplace(varName, ")", "_")
            varName := StrReplace(varName, "-", "_")
    
            ; Add a checkbox for the item
            Gui Add, Checkbox, x25 yp+30 v%varName%, %item% ;gsubmitChecklist

            ; Increment the inner loop index
            innerIndex++
        }
    }

; Add radio buttons for the "T&&Cs?" item in the "Pre SAP" section
; Gui Add, Text, tcs y+10, T&&Cs?
; Gui Add, Radio, x+5 gsubmitChecklist vtandcYes, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vtandcNa, N/A

; ---------- END PRE-SAP ----------