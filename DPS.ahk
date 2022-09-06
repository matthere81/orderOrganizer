#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
; SetWorkingDir, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs ; Ensures a consistent starting directory.
#SingleInstance Force
; #Include Order Organizer.ahk
#include <UIA_Interface>
#include <UIA_Browser>

;cUIA.FindFirstByName("Override Block").Click() ; Click Override
;text "Search Successfully Overridden." ; Wait for this text
;text "RESULTS" ; Click Results
;link "Generate" ; Click Generate
;header "Title" wait for this element
;text "No records found" ; No records found

; -------- Status Possibilites --------
; if Clear
;	Click generate
;	Wait for pdf - click


getDPSReports(cpq,customer,address,contact){
browserExe := "chrome.exe"
Run, %browserExe% --force-renderer-accessibility "https://hub.thermofisher.com/ip"
WinWaitActive, ahk_exe %browserExe%
cUIA := new UIA_Browser("ahk_exe " browserExe)
chrome := UIA.ElementFromHandle(browserExe, True)
WinWait GTC: Homepage - ONESOURCE Global Trade - Google Chrome
cUIA.WaitElementExistByName("DPS")
cUIA.FindFirstByName("DPS")
Sleep 1000
cUIA.FindFirstByName("DPS").Click() ; Click DPS
cUIA.WaitElementExistByName("DPS Search")
cUIA.FindFirstByName("DPS Search").Click() ; Click DPS Search
WinWait DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window

;-------- Input TENA-CPQ Number
referenceNumber := cUIA.FindByPath("1.2.2.1.1.1.9.2.2.1")
referenceNumber.Click()
referenceNumber.SetValue("TENA-CPQ-" . cpq)

;-------- Input Cust Name
dpsFirst := cUIA.FindByPath("1.2.2.1.1.1.8.2.2.1")
dpsFirst.Click()
dpsFirst.SetValue(customer)

;-------- Click Organization
organization := cUIA.FindByPath("1.2.2.1.1.1.8.3.2.1.1.1.1")
organization.Click()

;-------- Input Address
dpsAddress := cUIA.FindByPath("1.2.2.1.1.1.8.4.2.1")
dpsAddress.Click()
dpsAddress.SetValue(address)

cUIA.FindFirstByName("Search").Click() ; Click Search

WinWait DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
cUIA.WaitPageLoad("DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window")
; cUIA.FindFirstByNameAndType("Status", "text").Click() ; Click Status
cUIA.WaitElementExist
cUIA.WaitElementExist("Status")
MsgBox, Element exist
msgbox % cUIA.FindByPath("1.2.2.1.1.1.15.2.2.1").GetText()
; Clipboard := cUIA.DumpAll()



Return

;-------- Click Individual
individual := cUIA.FindByPath("1.2.2.1.1.1.8.3.2.1.1.2.1")
individual.Click()



; cUIA.WaitPageLoad("ONESOURCE Global Trade - Google Chrome")

Return

; dpsSecond := cUIA.FindByPath("1.2.2.1.1.1.8.3.2.1")
; dpsSecond.Click()
; dpsSecond.SetValue("Address")


Return ; - Put in until this script is working

cUIA.WaitPageLoad("DPS Search - ONESOURCE Global Trade - Google Chrome", 60000, 1000)
cUIA.WaitElementExistByName("ctl00$MainContent$txtLastDTSStatus")
cUIA.Select("ctl00$MainContent$txtLastDTSStatus")
Return
dpsStatus := cUIA.WaitElementExistByName("ctl00$MainContent$txtLastDTSStatus")
textValue := cUIA.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_TextControlTypeId) ; Create a TreeWalker to only look for ComboBox elements in the tree
TTW := cUIA.CreateTreeWalker(textValue) ; Create the TreeWalker with the Text condition
searchTW := TTW.GetNextSiblingElement(dpsStatus) ; Find the first ComboBox after the Google image
msgbox % searchTW.GetValue(searchTW)
; myField := cUIA.WaitElementExist("2.2.2.1.1.1.12.2.2.1")
; DPSResults()

; clipboard := cUIA.DumpAll()
Return
}

browserExe := "chrome.exe"
Run, %browserExe% --force-renderer-accessibility "https://hub.thermofisher.com/ip"
WinWaitActive, ahk_exe %browserExe%
cUIA := new UIA_Browser("ahk_exe " browserExe)
; cUIA.WaitTitleChange("DPS Search - ONESOURCE Global Trade - Google Chrome", 60000, 1000)
WinWait, DPS Search
MsgBox, here
Return


DPSResults(){
MsgBox, here
Loop
{   
	; No records found
    noRecords := cUIA.FindByFirstName("No Records Found")
    cUIA.FindFirstByName("Generate").Click() ; Click Search
	Break

	; No records found / Clear
	; CoordMode, Pixel, Screen
	; ImageSearch, FoundX, FoundY, 0, 0, 1920, 1080, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\clear.png
    ; If ErrorLevel = 0
	; {
	; 	Sleep, 200
	; 	Send, {tab 2}{enter}
	; 	Break
	; }
	
	; Blocked
	; CoordMode, Pixel, Window
	; ImageSearch, FoundX, FoundY, 959, 529, 1292, 617, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\blocked.png
	; If ErrorLevel = 0
	; {
	; 	MsgBox, 4, Found Records,Records WERE Found`nContinue?
	; 	If MsgBox No
	; 	{
	; 		Break
	; 	}
	; 	; Else
	; 	{
	; 		WinWaitActive, DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
	; 		MouseClick, Left, 899, 257,1,, D ; reset tab position to middle of screen
	; 		Send {tab 2}{enter}
	; 		Sleep 200
	; 		Send +{tab}{down 5}{enter}
	; 		Sleep 200
	; 		Send {tab 3}{Enter}
	; 		WinWaitActive, DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window
	; 		; MsgBox Is Tab -> Enter Correct here?
	; 		;Need search successfully overridden screenshot
	; 		; Send {tab}{Enter}
	; 		; Winwait here?
	; 		; Need wait for overridden screenshot here?
	; 		; Overridden
	; 		CoordMode, Pixel, Client
	; 		ImageSearch, FoundX, FoundY, 966, 487, 1241, 566, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\overridden.png
	; 		If ErrorLevel = 0
	; 		{
	; 			Sleep, 200
	; 			Send, {tab}{enter}
	; 		}
	; 		Break
	; 	}
	; }
	; ; Overridden
	; CoordMode, Pixel, Client
	; ImageSearch, FoundX, FoundY, 966, 487, 1241, 566, C:\Users\matthew.terbeek\AppData\Roaming\MacroCreator\Screenshots\overridden.png
	; If ErrorLevel = 0
	; {
	; 	Sleep, 200
	; 	Send, {tab 2}{enter}
	; 	Break
	; }
}
Return
}