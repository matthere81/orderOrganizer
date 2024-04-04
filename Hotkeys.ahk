
; ---- HOTKEY & HOT STRING LIST ----

; Esc to close the search GUI
SetTitleMatchMode 3
#IfWinActive, Order Organizer Search Results
Esc::Gui %MyGui%:Destroy
#IfWinActive
SetTitleMatchMode 2

; ---- END HOTKEY & HOT STRING LIST ----

Return