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

fields := ["SOT Line#", "Customer", "Customer Contact", "Sold To Account", "SO#", "Payment Terms"]
vars := ["sot", "customer", "contact", "soldTo", "soNumber", "terms"]

; Initialize the object with blank keys
values := {}
for index, key in vars
{
    values[key] := ""
}

; GUI Spacing Text & Edit Spacing
yTextField := 70
yEditField := 70

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
Gui Add, Button, gRestart y20, Restart ; Add a button that triggers the 'Restart' subroutine when clicked
Gui Add, Button, y20 gSaveToIni, &Save
Gui Add, Button, y20, &New PO/Order
Gui Add, Button, y20, Get Quote Info  ; Create a button


Gui Color, 79b8d1
Gui Font, S9, Segoe UI Semibold
; Gui Add, Button, x350 y20 w70 greadtheini, O&pen

; Gui Add, Button, x+25 w100 gClearFields, &Clear Fields


for index, field in fields
{
	Gui Add, Text, x50 y%yTextField%, % field
	yTextField += 50
	controlName := vars[index]
	Gui Add, Edit, yp+20 xp-2.5 v%controlName%
	yEditField += 20

	; Gui, Add, ControlType , Options, Text
}


Gui Show,w950 h700, Order Organizer ;SO# %soNumber%

#Include %A_ScriptDir%\Menu.ahk

Return

#Include %A_ScriptDir%\Functions.ahk
#Include %A_ScriptDir%\Search.ahk

