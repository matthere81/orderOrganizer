#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
workingDir := "C:\Users\" . A_UserName . "\OneDrive - Thermo Fisher Scientific\Documents\Order Organizer"
SetWorkingDir, workingDir ; Ensures a consistent starting directory.
#SingleInstance Force
#InstallKeybdHook	
#KeyHistory 50	
#include <Vis2>	
#include <UIA_Interface>	
#include <UIA_Browser>

; START START START START START START START START
;| ------------------------------------------ |
;| 			 GLOBAL VARIABLES  START		  |
;| ------------------------------------------ |
; START START START START START START START START

;  Create Order Organizer Path If It Doesn't Exist
if !FileExist(workingDir)
{
    FileCreateDir % workingDir
    SetWorkingDir % workingDir
}

; Include Icon
I_Icon = % workingDir . "\checkmark.ico"
IfExist, %I_Icon%
	Menu, Tray, Icon, %I_Icon%

global guiWidth := 925
global guiHeight := 510
global guiChecklistHeight := 765

; GUI Spacing Text & Edit Spacing
yTextField := 70
yEditField := 70
xCoordinate := 40
yCoordinate := 70
initialY := yCoordinate

; Create Order Database Path If It Doesn't Exist
myinipath := % workingDir . "\orderOrganizerDatabase"
if !FileExist(myinipath) {
    FileCreateDir, %myinipath%
}

if FileExist("C:\Users\" . A_UserName . "\Order Organizer")
{
	moveDatabase(myinipath)
}
	

;|------------------------------------------|
;|        Create Text and Edit Arrays       |
;|------------------------------------------|

fields := ["SOT Line#", "Customer", "Customer Contact", "Sold To Account", "Intercompany Entity", "Payment Terms", "SO#", "CPQ/Quote#", "System"
	, "CRD - Cust Req Date", "PO Date", "SAP Date", "Software S/N?", "PO#", "PO Value", "Tax", "Freight Cost", "Surcharge", "Total"
	, "Salesperson", "Sales Manager", "Sales Manager Code", "Sales Director", "Sales Director Code"
	, "Notes", "End User", "End User / Contact Phone", "End User / Contact Email", "End Use"]

vars := ["sot", "customer", "contact", "soldTo", "intercompanyEntity", "terms", "soNumber", "cpq", "system", "crd", "poDate", "sapDate", "serialNumber"
	, "po", "poValue", "tax", "freightCost", "surcharge", "totalCost", "salesperson", "salesManager", "managerCode", "salesDirector"
	, "directorCode", "notes", "endUser", "phone", "email", "endUse"]

; fieldsInfo := ""
; for index, field in fields
; {
; 	fieldsInfo .= "Index: " . index . ", Field: " . field . "`n"
; }
; MsgBox, % fieldsInfo

; Initialize the object with blank keys
values := {}
for index, key in vars
{
    values[key] := ""
}

;END END END END END END END END END END END
;| ------------------------------------------ |
;| 			 GLOBAL VARIABLES  END			  |
;| ------------------------------------------ |
;END END END END END END END END END END END

; GUI code goes here

Gui destroy
Gui +Resize +MinSize%guiWidth%x%guiHeight% +MaxSize%guiWidth%x%guiChecklistHeight% +HwndhGui
Gui Font
Gui Font, s12 w600 Italic cBlack, Tahoma
Gui Add, Text, x10 y30, _________________________________
Gui Add, Text, hWndhTxtOrderOrganizer23 x15 y20 w300 +Left, Order Organizer ; - SO# %soNumber%
Gui Font
Gui Add, Edit, vSearchTerm w200 y20 gSetSearchAsDefault ; Add an Edit field with 'Search for' as placeholder text
Gui Add, Button, Default gSearch y20, Search ; Add a button that triggers the 'Search' subroutine when clicked
Gui Add, Button, y20 gRestart, Restart ; Add a button that triggers the 'Restart' subroutine when clicked
Gui Add, Button, y20 gSaveToIni, &Save
Gui Add, Button, y20 gClearFields, &New PO/Order
Gui Add, Button, y20, Get Quote Info  ; Create a button
Gui Color, 79b8d1
Gui Font, S9, Segoe UI Semibold
Gui Add, StatusBar, vMyStatusBar -Theme

; ---- Loop through the fields and create the text and edit fields ----

for index, field in fields
{
	if (index = 8 or index = 14 or index = 20)
	{
		Gui Add, GroupBox, x+20 y95 w1 h335 ; vertical line
		xCoordinate += 175 ; Move to the next column
		yCoordinate := initialY ; Reset y-coordinate for the new column
	} 
	else if (index = 26)
	{
		Gui Add, GroupBox, x718 y95 w1 h225 ; vertical line
		xCoordinate += 175 ; Move to the next column
		yCoordinate := initialY ; Reset y-coordinate for the new column
	}

	Gui Add, Text, x%xCoordinate% y%yCoordinate%, % field . ":"
	yCoordinate += 50
	controlName := vars[index]

	if (field = "CRD - Cust Req Date" or field = "PO Date" or field = "SAP Date")
	{
		Gui Add, DateTime, yp+20 xp-2.5 w135 h22.5 v%controlName%,
		yCoordinate += 5
		Continue
	}

	if (field = "Notes")
	{
		Gui Add, Edit, yp+20 xp-2.5 w310 h90 v%controlName%
		yCoordinate += 60
		Continue
	}

	if (field = "End Use")
	{
		Gui Add, Edit, yp+20 xp-2.5 h90 v%controlName%
		yCoordinate += 60
		Continue
	}
	
	Gui Add, Edit, yp+20 xp-2.5 v%controlName% gEditChanged ;gAutoSave gFieldFocus
	yCoordinate += 5

}

; END ---- END - Loop through the fields and create the text and edit fields - END ---- END

; Gui Add, Progress, x0 y475 w950 h1 cRed ; Horizontal line
Gui Add, Text, x0 y475, ________________________________________________________________________________________________________________________________________________________________________________________________
; Gui, SubCommand [, Value1, Value2, Value3]
#Include ChecklistGUI.ahk
Gui Show, w%guiWidth% h%guiHeight%, Order Organizer ;SO# %soNumber%

; OnMessage(0x0201, "WM_LBUTTONDOWN") ; the formatting is weird and i don't know why
#Include %A_ScriptDir%\Menu.ahk
Return

#include %A_ScriptDir%\Hotkeys.ahk
#Include %A_ScriptDir%\Functions.ahk
#Include %A_ScriptDir%\Search.ahk

