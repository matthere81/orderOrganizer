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


; #Include %A_ScriptDir%\Hotstrings.ahk
; #Include %A_ScriptDir%\Shortcuts.ahk


; -------- GLOBAL VARIABLES -------- START

myinipath := "C:\Users\" . A_UserName . "\Order Organizer\Order Database"
if !FileExist(myinipath) {
    FileCreateDir, %myinipath%
}

fields := ["SOT Line#", "Customer", "Customer Contact", "Sold To Account", "SO#", "Payment Terms"]
vars := ["sot", "customer", "contact", "soldTo", "soNumber", "terms"]

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
Gui Add, Button, gRestart, Restart ; Add a button that triggers the 'Restart' subroutine when clicked
Gui Color, 79b8d1
Gui Font, S9, Segoe UI Semibold
; Gui Add, Button, x350 y20 w70 greadtheini, O&pen
; Gui Add, Button, x+25 w70 gSaveToIni, &Save
; Gui Add, Button, x+25 w125 grestartScript, &New PO or Reload
; Gui Add, Button, x+25 w100 gMyButton, Get Quote Info  ; Create a button
; Gui Add, Button, x+25 w100 gClearFields, &Clear Fields
; Gui Add, Edit, x+25 y22 w175 h20, Search

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

