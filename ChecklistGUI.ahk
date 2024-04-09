

;----------- START CHECKLISTS ---------------

Gui Add,Tab3,, Salesforce Checklist|SAP Checklist - Main Page|SAP Checklist - Inside The Order|SAP - Finalizing The Order
Gui Tab, 1

; ; Define the sections and their checklists
sections := ["Pre Salesforce", "Salesforce", "Pre SAP"]
checklists := [["Check TENA name on PO", "Quote number matches on PO && quote", "Payment terms match on PO && quote", "Prices match on PO && quote"
                , "Bill to && Ship to address on PO"],["Save PDF of Quote (Document Output Tab)", "Arrange quote lines (if needed) (Quote Details Tab)"
                , "Sold-To ID (Customer Details Tab)", "Check order type (Order Tab)", "Add PO# / Add PO Value / Upload PO (Order Tab)"]]
                ; , ["Order Notice Sent", "Entered In SOT", "T&&Cs?"]]

; Loop over each section
Loop, % sections.MaxIndex()
    {
        ; Calculate y-coordinate based on section index
        yCoord := 600 ; Adjust multiplier as needed for spacing
        xCoord := 10

        ; Add a group box for the section
        Gui Add, GroupBox, y%yCoord% h175 w250, % sections[A_Index]
    
        index := A_Index
        ; Loop over each item in the checklist
        Loop, % checklists[A_Index].MaxIndex()
        {
            ; Calculate y-coordinate for the checkbox based on item index
            ; checkboxYCoord := ys+10 ;yCoord + 25 ; + (A_Index * 30) ; Adjust multiplier as needed for spacing
            
            ; Get the checklist item
            item := checklists[index,A_Index] ;,[A_Index[A_Index]] ;,%index%]
            ; MsgBox % index . " " . A_Index
            ; MsgBox % item

            ; Create a sanitized version of the item to use as the variable name
            varName := StrReplace(item, " ", "_")
            varName := StrReplace(varName, "&", "_")
            varName := StrReplace(varName, "(", "_")
            varName := StrReplace(varName, ")", "_")
            varName := StrReplace(varName, "-", "_")
            varName := StrReplace(varName, "/", "_")
    
            ; Add a checkbox for the item
            ; Gui Add, Checkbox, x10 y%yCoord% + 10 v%varName%, %item% ;y%checkboxYCoord% v%varName%, %item% ;gsubmitChecklist

        }
    }

; Add radio buttons for the "T&&Cs?" item in the "Pre SAP" section
; Gui Add, Text, tcs y+10, T&&Cs?
; Gui Add, Radio, x+5 gsubmitChecklist vtandcYes, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vtandcNa, N/A

; ---------- END PRE-SAP ----------