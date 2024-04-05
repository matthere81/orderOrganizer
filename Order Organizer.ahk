#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir, C:\Users\%A_UserName%\Order Organizer ; Ensures a consistent starting directory.
#SingleInstance Force
#InstallKeybdHook	
#KeyHistory 50	
#include <Vis2>	
#include <UIA_Interface>	
#include <UIA_Browser>


;  Create Order Organizer Path If It Doesn't Exist
if !FileExist("C:\Users\" . A_UserName . "\Order Organizer") {
    FileCreateDir, A_WorkingDir . "\Order Organizer"
    SetWorkingDir, A_WorkingDir . "\Order Organizer"
}

; Include Icon
; FileInstall, C:\Users\A_UserName\Order Organizer\list_check_checklist_checkmark_icon_181579.ico, A_WorkingDir, 1
I_Icon = C:\Users\%A_UserName%\Order Organizer\checkmark.ico
IfExist, %I_Icon%
	Menu, Tray, Icon, %I_Icon%

; -------- GLOBAL VARIABLES -------- START

myinipath := "C:\Users\" . A_UserName . "\Order Organizer\Order Database"
if !FileExist(myinipath) {
    FileCreateDir, %myinipath%
}

fields := ["SOT Line#", "Customer", "Customer Contact", "Sold To Account", "Intercompany Entity", "SO#", "Payment Terms", "CPQ/Quote#", "System"
	, "CRD - Cust Req Date", "PO Date", "SAP Date", "PO#", "PO Value", "Tax", "Freight Cost", "Surcharge", "Total"
	, "Salesperson", "Sales Manager", "Sales Manager Code", "Sales Director", "Sales Director Code", "Software S/N?"
	, "Notes", "End User", "End User / Contact Phone", "End User / Contact Email", "End Use"]

vars := ["sot", "customer", "contact", "soldTo", "intercompanyEntity", "soNumber", "terms", "cpq", "system", "crd", "poDate", "sapDate", "po"
	, "poValue", "tax", "freightCost", "surcharge", "totalCost", "salesperson", "salesManager", "managerCode", "salesDirector"
	, "directorCode", "serialNumber", "notes", "endUser", "phone", "email", "endUse"]

; Initialize the object with blank keys
values := {}
for index, key in vars
{
    values[key] := ""
}

; GUI Spacing Text & Edit Spacing
yTextField := 70
yEditField := 70
xCoordinate := 40
yCoordinate := 70
initialY := yCoordinate


; -------- GLOBAL VARIABLES -------- END

; GUI code goes here

Gui destroy
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


; ---- Loop through the fields and create the text and edit fields ----
for index, field in fields
{
	if (index = 7 or index = 13 or index = 19 or index = 25)
	{
		Gui Add, GroupBox, x+20 y95 w1 h285 ; vertical line
		xCoordinate += 175 ; Move to the next column
		yCoordinate := initialY ; Reset y-coordinate for the new column
	}

	Gui Add, Text, x%xCoordinate% y%yCoordinate%, % field . ":"
	yCoordinate += 50
	controlName := vars[index]

	if (field = "CRD - Cust Req Date" or field = "PO Date" or field = "SAP Date")
	{
		Gui Add, DateTime, yp+20 xp-2.5 w135 h20 v%controlName%,
		yCoordinate += 5
		Continue
	}

	if (field = "Notes")
	{
		Gui Add, Edit, yp+20 xp-2.5 w135 h90 v%controlName%
		yCoordinate += 60
		Continue
	}

	Gui Add, Edit, yp+20 xp-2.5 v%controlName%
	yCoordinate += 5
}

; END ---- END - Loop through the fields and create the text and edit fields - END ---- END


Gui Show,w950 h700, Order Organizer ;SO# %soNumber%

#Include %A_ScriptDir%\Menu.ahk
Return

#include %A_ScriptDir%\Hotkeys.ahk
#Include %A_ScriptDir%\Functions.ahk
#Include %A_ScriptDir%\Search.ahk

