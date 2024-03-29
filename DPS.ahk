#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
; SetWorkingDir, C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Documents\Order Docs\SO Docs ; Ensures a consistent starting directory.
#SingleInstance Force
; #Include Order Organizer.ahk
#include <UIA_Interface>
#include <UIA_Browser>
#include <UIA_Constants>

SetTitleMatchMode 2

browserExe := "chrome.exe"
; Run, %browserExe% --force-renderer-accessibility "https://tfs-3.lightning.force.com/lightning/r/Opportunity/0064z00002CgNdFAAV/view"
Run, %browserExe% --force-renderer-accessibility "https://tfs-3--c.vf.force.com/apex/porder_createWIN_p?oppId=0064z00002CgNdFAAV&"
WinWaitActive, ahk_exe %browserExe%
cUIA := new UIA_Browser("ahk_exe " browserExe)
Sleep 2000
opportunity := "Orbitrap Exploris 240 IC - NA.schen - Ionis Pharmaceuticals - Hans Gaus - PN-1795317"

cUIA.WaitElementExistByName("WIN Form")
winFormLink := cUIA.GetCurrentURL()
winForm := cUIA.FindFirstByName("WIN Form")
pattern := "Opportunity/(.*)/view"
RegExMatch(winFormLink, pattern, match)
match1 := Trim(match1)
winFormLink := cUIA.SetURL("https://tfs-3--c.vf.force.com/apex/porder_createWIN_p?oppId=" . match1 . "&", True)
cUIA.WaitPageLoad(winFormLink)

endUser := cUIA.FindFirstByName("End User Contact Name")
endUser := endUser.FindByPath("+1")

phoneNumber := cUIA.FindFirstByName("End User Phone#")
phoneNumber := phoneNumber.FindByPath("+1")

email := cUIA.FindFirstByName("End User Email")
email := email.FindByPath("+1")

endUse := cUIA.FindFirstBy("AutomationID=thepage\:theform\:j_id109")

endUser := endUser.Name
phoneNumber := phoneNumber.Name
email := email.Name
endUse := endUse.Value


MsgBox % endUser . "`n" . phoneNumber . "`n" . email . "`n" . endUse
; suggestions := cUIA.FindFirstByName("Suggestions")
; suggestions.SetFocus()
; suggestions := cUIA.FindFirstByNameAndType(opportunity, "option")
; suggestions.Highlight()
Return

; -------- Status Possibilites --------
; if Clear
;	Click generate
;	Wait for pdf - click

; -------- DO THE DUMP --------
; Clipboard := cUIA.DumpAll()

; -------- Hardcode Testing Vars --------
cpq := "00535570"
customer := "United Salt Corporation"
contact := "Kevin Gardner"
address := "7901 FM 1405"

; getDPSReports(cpq,customer,address,contact){
browserExe := "chrome.exe"
Run, %browserExe% --force-renderer-accessibility "https://hub.thermofisher.com/ip"
WinWaitActive, ahk_exe %browserExe%
cUIA := new UIA_Browser("ahk_exe " browserExe)
WinWait GTC: Homepage - ONESOURCE Global Trade - Google Chrome
Sleep 1000

; -------- Click DPS and DPS Search on homepage --------
cUIA.WaitElementExistByName("DPS")
cUIA.FindFirstByName("DPS")
cUIA.FindFirstByName("DPS").Click()
cUIA.WaitElementExistByName("DPS Search")
cUIA.FindFirstByName("DPS Search").Click()

WinWaitActive DPS Search - ONESOURCE Global Trade - Google Chrome, Chrome Legacy Window

defaultClick := "left","1, 500"

; -------- Set paths for text entries --------
searchName := cUIA.FindByPath("1.1.2.1.1.1.6.2.2.1")
organization := cUIA.FindByPath("1.1.2.1.1.1.6.3.2.1.1.1.1")
individual := cUIA.FindByPath("1.1.2.1.1.1.6.3.2.1.1.2.1")
addressField := cUIA.FindByPath("1.1.2.1.1.1.6.4.2.1")
tenaCpq := cUIA.FindByPath("1.1.2.1.1.1.7.2.2.1")
screenNow := cUIA.FindFirstByName("Screen now")
; status := cUIA.FindByPath("1.1.2.1.1.1.13.3")
; status := cUIA.FindFirstByName("Status").FindByPath("+1")
status := cUIA.FindByPath("1.1.2.1.1.1.13.2.2.1")


cUIA.WaitElementExistByName("Postal Code")
cUIA.WaitElementExistByName("Search Name")
searchName.SetFocus()

while (!searchName.HasKeyboardFocus = 1)
{
	Sleep 200
	searchName.SetFocus()
	Break
}

; -------- Input Organization Name -------- ;
searchName.SetValue(customer)
Sleep 250

; -------- Click Organization -------- ;
organization.Click(defaultClick)
Sleep 250
addressField.SetValue(address)
Sleep 250

;-------- Input TENA-CPQ Number -------- ;
tenaCpq.SetValue("TENA-CPQ-" . cpq)
Sleep 250

; -------- Input Address -------- ;
screenNow.Click(defaultClick)

cUIA.WaitElementExistByName("Override Block")

; -------- GET TEXT FROM FIELD -------- ;
textPattern := status.GetCurrentPatternAs("Text")
dpsStatus := textPattern.DocumentRange


if dpsStatus = "Overridden"
{
	MsgBox % "Status is " . dpsStatus
}
else if dpsStatus = "Blocked"
{
	MsgBox % "Status is " . dpsStatus . "`nCheck results"
	
	; -------- Finds the element with number of records and returns a string -------- ;
	recordsFound := cUIA.FindByPath("1.1.2.1.1.1.9.1.1.2") ;.Highlight()
	recordsFound := recordsFound.CurrentName
	recordsFound := RegExMatch(recordsFound, "(\d)")
	
	if recordsFound < 25
	{
		; -------- VIEW ALL RECORDS -------- :
		
		Send ^f
		Sleep 200
		Send customer
		numberFoundOnPage := cUIA.FindFirstByType("50017")
		numberFoundOnPage := numberFoundOnPage.CurrentName
		numberFoundOnPage := RegExMatch(numberFoundOnPage, "(?:.*(\d$))")
		
		If numberFoundOnPage <=2
		{
			cUIA.FindFirstByName("Override Block").Click()
			Send +{tab}{down 5}{Enter}
			Sleep 200
			Send {tab 3}
			Sleep 200
			Send {enter}

			cUIA.WaitElementExist("Search Successfully Overridden.")
			cUIA.FindFirstByName("RESULTS").Click()
			cUIA.FindFirstByName("Screening Report").Click()
			cUIA.WaitElementExist("Title")
			cUIA.FindFirstByName("Title").Click("left", "2, 500")
			Sleep 1000
			Send {Tab}
			Send {Enter}
			; Wait for DTSSearch Results
		}
		Else
			Return
	} 
	Else 
	{
		MsgBox % recordsFound . " from else"

		; -------- VIEW ALL RECORDS -------- :

		; -------- Get all elements with name of select (returns array) -------- ;
		dropdowns := cUIA.FindAllByName("select")

		; -------- Get last element in array -------- ;
		lastElement := dropdowns[dropdowns.Length()]
		lastElement.Click()
		Sleep 200
		Send {down 2}
		Sleep 200
		Send {Enter}
		; -------- END VIEW ALL RECORDS -------- :
	}
	; ---- END Finds the element with number of records and returns a string -------- ;
}
else if dpsStatus = "Clear"
	cUIA.FindFirstByName("No records found.")
	; ^p to print screen 
Return

; -------- DOES THE DUMP -------- ;
; Clipboard := cUIA.DumpAll()


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
; }

; browserExe := "chrome.exe"
; Run, %browserExe% --force-renderer-accessibility "https://hub.thermofisher.com/ip"
; WinWaitActive, ahk_exe %browserExe%
; cUIA := new UIA_Browser("ahk_exe " browserExe)
; cUIA.WaitPageLoad("GTC: Homepage - ONESOURCE Global Trade - Google Chrome",,1000)

; dlLink := cUIA.FindFirstByName("Download")
; MsgBox, % cUIA.DumpAll()
; cUIA.WaitTitleChange("DPS Search - ONESOURCE Global Trade - Google Chrome", 60000, 1000)
; WinWait, DPS Search
; WinWaitActive GTC: Homepage - ONESOURCE Global Trade - Google Chrome
Sleep 3000
; MsgBox, here
; Clipboard := cUIA.DumpAll()
; ReturnW


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