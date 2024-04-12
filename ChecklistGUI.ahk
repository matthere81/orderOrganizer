

;----------- START CHECKLISTS ---------------

Gui Add,Tab3,, Salesforce Checklist|SAP Checklist - Main Page|SAP Checklist - Inside The Order|SAP - Finalizing The Order
Gui Tab, 1

;|------------------------------------------|
;|                                          |
;| Define the section 1 and its checklists  |
;|                                          |
;|------------------------------------------|

section1 := ["Pre Salesforce", "Salesforce", "Pre SAP"]
checklists1 := [["Check TENA name on PO", "Quote number matches on PO && quote", "Payment terms match on PO && quote", "Prices match on PO && quote"
                , "Bill to && Ship to address on PO"], ["Save PDF of Quote (Document Output Tab)", "Arrange quote lines (if needed) (Quote Details Tab)"
                , "Sold-To ID (Customer Details Tab)", "Check order type (Order Tab)", "Add PO# / Add PO Value / Upload PO (Order Tab)"]
                , ["Order Notice Sent", "Entered In SOT", "T&&Cs?"]]

; Loop over each section
Loop, % section1.MaxIndex()
{
    index := A_Index
    ; Call the subroutine with the current section index, section and checklist
    ; Func("AddSection").Call(section1[A_Index], checklists1[A_Index])
    AddSection(section1[A_Index], checklists1)
}


;|------------------------------------------|
;|                                          |
;| Define the section 2 and its checklists  |
;|                                          |
;|------------------------------------------|

section2 :=["SAP Attachments", "Main Sales Tab"]
checklists2 := [["PO", "Quote", "Order Notice", "WIN Form", "Merge Report"], ["CRD Date", "First Date Lines Updated", "Incoterms"]]

; ; Loop over each section
; Loop, % section2.MaxIndex()
; {
;     ; Calculate y-coordinate based on section index
;     yCoord := 575 ; Adjust multiplier as needed for spacing
;     xCoord := 10 + (A_Index - 1) * 300 ; Adjust multiplier as needed for spacing

;     ; Add a group box for the section
;     Gui Add, GroupBox, x%xCoord% y%yCoord% h175 w280, % section1[A_Index]

;     checkboxYCoord := 595 ;+ (A_Index - 1) * 20 ; Adjust multiplier as needed for spacing

;     index := A_Index
;     ; Loop over each item in the checklist
;     Loop, % checklists1[A_Index].MaxIndex()
;     {
;         ; Calculate x-coordinate for the checkbox based on item index
;         checkboxXCoord := % xCoord + 10 ;+ (A_Index - 1) ; * 15 ; Adjust multiplier as needed for spacing
        
;         ; Get the checklist item
;         item := checklists1[index,A_Index] ;,[A_Index[A_Index]] ;,%index%]

;         ; Create a sanitized version of the item to use as the variable name
;         varName := StrReplace(item, " ", "_")
;         varName := StrReplace(varName, "&", "_")
;         varName := StrReplace(varName, "(", "_")
;         varName := StrReplace(varName, ")", "_")
;         varName := StrReplace(varName, "-", "_")
;         varName := StrReplace(varName, "/", "_")
;         varName := StrReplace(varName, "?", "_")

;         ; If the item text is longer than 45 characters, split it into two lines
;         if (StrLen(item) > 40)
;         {
;             ; Find the position of the next whitespace character after the 45th character
;             pos := InStr(item, " ", false, 41)
;             ; If a whitespace character was found, split the item text at that position
;             if (pos > 0)
;             {
;                 line1 := SubStr(item, 1, pos - 1)
;                 line2 := SubStr(item, pos + 1)
;             }
;             else
;             {
;                 ; If no whitespace character was found, split the item text at the 45th character
;                 line1 := SubStr(item, 1, 45)
;                 line2 := SubStr(item, 46)
;             }
        
;             Gui Add, Checkbox, x%checkboxXCoord% y%checkboxYCoord% v%varName%, %line1%`n%line2%
;             ; Increase the y-coordinate for the next checkbox by an additional amount to account for the extra line
;             checkboxYCoord := checkboxYCoord + 30 ; Adjust as needed for spacing
;         }
;         else
;         {
;             Gui Add, Checkbox, x%checkboxXCoord% y%checkboxYCoord% v%varName%, %item%
;             ; Increase the y-coordinate for the next checkbox
;             checkboxYCoord := checkboxYCoord + 20 ; Adjust as needed for spacing
;         }
;     }
; }

; Gui Tab, 2

; Add radio buttons for the "T&&Cs?" item in the "Pre SAP" section
; Gui Add, Text, tcs y+10, T&&Cs?
; Gui Add, Radio, x+5 gsubmitChecklist vtandcYes, Yes
; Gui Add, Radio, x+5 gsubmitChecklist vtandcNa, N/A

; ---------- END PRE-SAP ----------