#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

; Path to the Excel workbook
workbookPath := "C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\General\Training Docs\Sales List fed 12.xlsx"
; workbookPath := "C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Desktop\Book1.xlsx"
salesWorkBook := "Sales List fed 12.xlsx"
lastModifiedIni := "C:\Users\matthew.terbeek\OneDrive - Thermo Fisher Scientific\Desktop\Auto Hot Key Scripts\orderOrganizer\lastModified.ini"

checkLastModified(workbookPath,lastModifiedIni)

; Disables ComObj errors
ComObjError(false)

xl := ComObjActive("Excel.Application")
if (xl != "")
{
    wb := xl.Workbooks(salesWorkBook)
    if (wb != "")
    {
        MsgBox, The workbook is open.
    }
    else
    {
        MsgBox, The workbook is not open.
    }
}
else
{
    ; the Excel application is not open, so create the COM object
	xl := ComObjCreate("Excel.Application")
	; open the workbook
	wb := xl.Workbooks.Open(workbookPath)
}

;-------- LIST OUT WORKSHEET NAMES --------;
; Select the worksheet named "Digital Sales"
ws1 := wb.Worksheets[1] ; - West Denise Schwartz
ws2 := wb.Worksheets[2] ; - East & Canada Maroun
ws3 := wb.Worksheets[3] ; - Digital
ws4 := wb.Worksheets[4] ; - IOMS Sales
ws5 := wb.Worksheets[5] ; - WiAS Team
;-------- END LIST OUT WORKSHEET NAMES END --------;


; Identify sales regions
ws1Regions := ["AMER-GCM", "AMER-MWM", "AMER-MOM", "AMER-NWM", "AMER-SWM"]

; Identify sales manager for that region
	; Loop through each ws1Regions and find individual region
	; Identify the next colored cell (the sales manager)
	; Get sales people in that range and assign them to the group



; Get the range of cells that contain data
usedRange := ws1.UsedRange

rowCount := usedRange.Rows.Count
colCount := usedRange.Columns.Count
; xlWhole := 1

; Find the first cell containing the search value
; searchValue := ws1Regions.1


; Loop through each value in the array
for eachRegion in ws1Regions
{
	foundCell := usedRange.Find(ws1Regions[eachRegion])
	currentRegion := ws1Regions[eachRegion]
	; MsgBox, % ws1Regions[eachRegion]
	
	if (foundCell <> "")
	{
		; Get the address of the found cell
		foundAddress := foundCell.Address

		; Display the address in a message box
		MsgBox, % foundCell.Value
		MsgBox, The value was found in cell %foundAddress%.
	}
	else
	{
		MsgBox, % ws1Regions[eachRegion]
	}

    ; Search for the region in the worksheet
    ; cell := usedRange.Find(currentRegion)

	; MsgBox % cell
}

wb.Close()
MsgBox, end




Return

; Disables ComObj errors
ComObjError(false)

; Check if Excel is already running
xl := ComObjActive("Excel.Application")

; check if Excel is already open and create the COM object
If (xl := ComObjActive("Excel.Application"))
{
    ; check if the workbook is already open
    wb := xl.Workbooks.Item(Book1.xlsx)
	
	if (wb = 0)
    {
	}
	Else
	{
		
    	MsgBox, The workbook has been opened.
	}
}

ws := wb.Worksheets("Sheet1")

; Read the value of cell A1
value := ws.Range("A1").Value

MsgBox, % value

wb.Close()
MsgBox, end
Return

; Check if workbook is open
; wb := xl.Workbooks(workbookPath)
; MsgBox % wb
; if (wb is not equal to "") {
;     MsgBox, % "Workbook is open"
; } else {
;     MsgBox, % "Workbook is not open"
; }
; }

; Check if workbook is open
; wb := xl.Workbooks(workbookPath)

; Open the workbook
; wb := xl.Workbooks.Open(workbookPath)

MsgBox, end
xl.Quit()
Return
;-------- LIST OUT WORKSHEET NAMES --------;
; Select the worksheet named "Digital Sales"
ws1 := wb.Worksheets[1] ; - West Denise Schwartz
ws2 := wb.Worksheets[2] ; - East & Canada Maroun
ws3 := wb.Worksheets[3] ; - Digital
ws4 := wb.Worksheets[4] ; - IOMS Sales
ws5 := wb.Worksheets[5] ; - WiAS Team
;-------- END LIST OUT WORKSHEET NAMES END --------;


; Identify sales regions
ws1Regions := ["AMER-GCM", "AMER-MWM", "AMER-MOM", "AMER-NWM", "AMER-SWM"]

; Identify sales manager for that region
	; Loop through each ws1Regions and find individual region
	; Identify the next colored cell (the sales manager)
	; Get sales people in that range and assign them to the group



; Get the range of cells that contain data
usedRange := ws1.UsedRange

rowCount := usedRange.Rows.Count
colCount := usedRange.Columns.Count
; xlWhole := 1

; Find the first cell containing the search value
; searchValue := ws1Regions.1


; Loop through each value in the array
for eachRegion in ws1Regions
{
	foundCell := usedRange.Find(ws1Regions[eachRegion])
	currentRegion := ws1Regions[eachRegion]
	; MsgBox, % ws1Regions[eachRegion]
	
	if (foundCell <> "")
	{
		; Get the address of the found cell
		foundAddress := foundCell.Address

		; Display the address in a message box
		MsgBox, % foundCell.Value
		MsgBox, The value was found in cell %foundAddress%.
	}
	else
	{
		MsgBox, % ws1Regions[eachRegion]
	}

    ; Search for the region in the worksheet
    ; cell := usedRange.Find(currentRegion)

	; MsgBox % cell
}

; Check if the search value was found
; if (foundCell)
; {
;     ; Get the address of the found cell
;     foundCellAddress := foundCell.Address
    
;     ; Get the value in the found cell
;     foundCellValue := foundCell.Value
    
;     ; Display the results
;     MsgBox, Found cell %foundCellAddress% with value %foundCellValue%.
; }
; else
; {
;     MsgBox, Could not find %searchValue%.
; }

; Close the workbook and quit Excel
wb.Close()

MsgBox excel closed

Return

ws1ManagerA := []

counter := 0

; Loop through all cells in column A
for row in ws1.Range("A1:A" rowCount)
{
    cell := row.Cells[1]
    bgColor := cell.Interior.Color
	whiteRGB := 16777215
    if (bgColor <> whiteRGB) ; Check if the background color is not white
    {
		if (bgColor && counter <= 1)
		{
			ws1ManagerA.Push(cell.Address)
			counter++
			; MsgBox, % cell.Address
			if counter = 1
			{
				startingCell := cell.Address
			}
			if counter = 2
			{
				lastCell := cell.Address
				manager := ws1ManagerA.Push(cell.Value)
			}
		}
	}
}

ws1ManagerA.RemoveAt(1,2)

; ws1ManagerA.push(manager)
; MsgBox, % startingCell . " " . lastCell

; Display all the values in the array
for Index, Value in ws1ManagerA
	Display .= Value . "`n"
MsgBox % Display

; When finished, close the workbook and quit Excel
wb.Close()
xl.Quit()


Return


;-------- FUNCTIONS --------;

checkLastModified(workbookPath,lastModifiedIni)
{
; Get the modification time of the file
FileGetTime, modTime, %workbookPath%, M

; Read the last checked time from a file
FileRead, lastCheckTime, lastModifiedIni

Return
; If the last check time is more than a week ago, check for changes
if (A_Now - lastCheckTime >= 60) ;7 * 24 * 60 * 60)
{
    ; Update the last checked time in the file
    FileDelete, lastModifiedIni
    FileAppend, %A_Now%, lastModifiedIni

    ; Get the modification time of the file again
    FileGetTime, modTime, %workbookPath%, M

    ; Compare the modification time to the previous time
    if (modTime <> prevModTime)
    {
        ; Create an Outlook application object
		Outlook := ComObjCreate("Outlook.Application")

		; Create an email item
		Email := Outlook.CreateItem(0)

		; Set the email properties
		Email.To := "matthew.terbeek@thermofisher.com"
		Email.Subject := "Sales Change"
		Email.Body := "This is a test email sent from AutoHotkey."


		; Send the email
		Email.Send()

		; Release the objects
		Email := ""
		Outlook := ""
        prevModTime := modTime
    }
}
}




Return









Data := {}

; Get all cell values in column A and put in an array
for Cell in ws1.UsedRange.Cells
	Data.Push(Cell.Value)

; Display all the values in the array
for Index, Value in Data
	Display .= Index "`t" Value ;"`n"
MsgBox % Display

; When finished, close the workbook and quit Excel
wb.Close()
xl.Quit()


Return

/*

; Initialize an empty array to store the cell values
values := []

; Loop through the range and collect the cell values in the array
for row in range
{
    cellValue := row.Value
    values.Push(cellValue)
}

; Initialize an empty string to hold the message
valuesString := ""


; Concatenate each value and add a newline character
for index, value in values
{
	valuesString .= value . "`n"
}

; Display the message box
msgBox % valuesString

; Display the collected cell values in a message box
; MsgBox % "The cell values are: " values.Join(", ")
; MsgBox % values.Join(", ")

; When finished, close the workbook and quit Excel
wb.Close()
xl.Quit()


Return

/*

salesPeople := "|Justin Carder|Robin Sutka|Fred Simpson|Rhonda Oesterle|Mitch Lazaro|Tucker Lincoln|Jawad Pashmi|Julie Sawicki|Mike Hughes|"
. "Steve Boyanoski|Bob Riggs|Chuck Costanza|Navette Shirakawa|Stephanie Koczur|Mark Krigbaum|Jon Needels|Bill Balsanek|Brent Boyle|Andrew Clark"
. "|Kevin Clodfelter|Gabriel Mendez|Karl Kastner|Michael Burnett|Jerry Pappas|Nick Duczak|Steven Danielson|Nick Hubbard (Nik)|"
. "Samantha Stikeleather|Drew Smillie|Jeff Weller|Jerry Holycross|Theresa Borio|Dan Ciminelli|Cynthia (Cindy) Spittler|Gwyn Trojan"
. "|Joel Stradtner|Don Rathbauer|Hillary Tennant|Melissa Chandler|Douglas Sears|Rashila Patel|Brian Thompson|Larry Bellan|Donna Zwirner"
. "|Kristen Luttner|Helen Sun|May Chou|Haris Dzaferbegovic|Brian Dowe|Mark Woodworth|Susan Bird|Giovanni Pallante|Alicia Arias"
. "|Dominique Figueroa|Jonathan McNally|Murray Fryman|Yan Chen|Jie Qian|Joe Bernholz|David Kage|David Scott|Todd Stoner|John Bailey|"
. "Katianna Pihakari|Jonathan Ferguson|Aeron Avakian|Luke Marty|Alexander James|Timothy Johnson|Yuriy Dunayevskiy|Susan Gelman"
. "|Cari Randles|Shijun (Simon) Sheng|Sean Bennett|Nelson Huang|Lorraine Foglio|Gerald Koncar|Lauren Fischer|Brian Luckenbill"
. "|Amy Allgower|Brandon Markle|Crystal Flowers|Douglas McDowell|Dante Bencivengo|Dana Stradtner|Justin Chang|Kate Lincoln|"
. "Angelito Nepomuceno|Patrick Bohman|Kristin Roberts|John Venesky|Sarah Jackson|Daniel Quinn|Eric Norviel|Lisa Kasper|Karla Esparza|"
. "Loris Fossir|Russ Constantineau|David Kusel|Taylor Graham|Jerome Johemko|Christie Baldizar|Christina Guintu|Ross Milam|Sunny Chen|David Hill"
. "|Elaine Miller|Gary Scharrer|Valerie Bruner|Laura Howell|Cecilia Snyder|Tatiana Valle Melendez|Annie Cantelmo|Sitara Chauhan|Doug Meinhart"
. "|Tori Milioni|Bob Myers|Paulette Parker|Jane-Marie Kowalski|Ronsar Eid|Brian Ridley|Krystina Simms|"

salesManagers := "|Anjou Keller|Joe Hewitt|Zee Nadjie|Doug McCormack|Natalie Foels|Tonya Second|Lou Gavino|Christopher Crafts|Joe McFadden|John Butler|Richard Klein|Ray Chen|Randy Porch"

salesDirectors := "|Denise Schwartz|Joann Purkerson|Maroun El Khoury|Jimmy Yuk|N/A"

salesCodes := "|201020|202375|96715|1261|98866|96695|96654|202625|202006|1076|95410|202610|1026|1042|202756|202611|1041|N/A"

kellerDropDown:
	GuiControl, Choose, salesManager, Anjou Keller
	GuiControl, ChooseString, managerCode, 202375
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

mccormackDropDown:
	GuiControl, Choose, salesManager, Doug McCormack
	GuiControl, ChooseString, managerCode, 1261
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

hewittDropDown:
	GuiControl, Choose, salesManager, Joe Hewitt
	GuiControl, ChooseString, managerCode, 98866
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

foelsDropDown:
	GuiControl, Choose, salesManager, Natalie Foels
	GuiControl, ChooseString, managerCode, 96715
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

secondDropDown:
	GuiControl, Choose, salesManager, Tonya Second
	GuiControl, ChooseString, managerCode, 95410
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

mcfaddenDropDown:
	GuiControl, Choose, salesManager, Joe McFadden
	GuiControl, ChooseString, managerCode, 202610
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, N/A
	GuiControl, ChooseString, directorCode, N/A
	Gui Submit, NoHide
return

butlerDropDown:
	GuiControl, Choose, salesManager, John Butler
	GuiControl, ChooseString, managerCode, 1026
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

kleinDropDown:
	GuiControl, Choose, salesManager, Richard Klein
	GuiControl, ChooseString, managerCode, 1042
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

nadjieDropDown:
	GuiControl, Choose, salesManager, Zee Nadjie
	GuiControl, ChooseString, managerCode, 96695
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

chenDropDown:
	GuiControl, Choose, salesManager, Ray Chen
	GuiControl, ChooseString, managerCode, 202756
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Jimmy Yuk
	GuiControl, ChooseString, directorCode, 202611
	Gui Submit, NoHide
return

porchDropDown:
	GuiControl, Choose, salesManager, Randy Porch
	GuiControl, ChooseString, managerCode, 1041
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Maroun El Khoury
	GuiControl, ChooseString, directorCode, 1076
	Gui Submit, NoHide
return

tollstrupDropDown:
	GuiControl, Choose, salesManager, Darren Tollstrup
	GuiControl, ChooseString, managerCode, 200320
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Sylveer Bergs
	GuiControl, ChooseString, directorCode, 203185
	Gui Submit, NoHide
return

gavinoDropDown:
	GuiControl, Choose, salesManager, Lou Gavino
	GuiControl, ChooseString, managerCode, 96654
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 202375
	Gui Submit, NoHide
return

craftsDropDown:
	GuiControl, Choose, salesManager, Christopher Crafts
	GuiControl, ChooseString, managerCode, 202625
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, Denise Schwartz
	GuiControl, ChooseString, directorCode, 201020
	Gui Submit, NoHide
return

findSales:
;=============== KELLER ===================
if salesPerson = Julie Sawicki
	gosub, kellerDropDown
if salesPerson = Justin Carder 
	gosub, kellerDropDown
if salesPerson = Brent Boyle
	gosub, kellerDropDown
if salesPerson = Luke Marty
	gosub, kellerDropDown
if salesPerson = Aeron Avakian 
	gosub, kellerDropDown
if salesPerson = Jon Needels
	gosub, kellerDropDown
if salesPerson = Brian Dowe
	gosub, kellerDropDown
if salesPerson = Lisa Kasper
	gosub, kellerDropDown
if salesPerson = John Bailey
	gosub, kellerDropDown
if salesPerson = Aeron Avakian
	gosub, kellerDropDown
if salesPerson = Luke Marty
	gosub, kellerDropDown
if salesPerson = Brandon Markle
	gosub, kellerDropDown
;=============== END KELLER ===================

;============ NADJIE ================
if salesPerson = Jawad Pashmi
	gosub, nadjieDropDown
if salesPerson = Navette Shirakawa
	gosub, nadjieDropDown
if salesPerson = Alicia Arias
	gosub, nadjieDropDown
if salesPerson = Gabriel Mendez
	gosub, nadjieDropDown
if salesPerson = Rhonda Oesterle
	gosub, nadjieDropDown
if salesPerson = Alexander James
	gosub, nadjieDropDown
if salesPerson = Shijun (Simon) Sheng
	gosub, nadjieDropDown
if salesPerson = Brian Luckenbill
	gosub, nadjieDropDown
if salesPerson = Amy Allgower
	gosub, nadjieDropDown
if salesPerson = Michael Burnett
	gosub, nadjieDropDown
if salesPerson = Karla Esparza
	gosub, nadjieDropDown
;============ END NADJIE ================

;========== DOUG MCCORMACK =================
if salesPerson = Mark Krigbaum
	gosub, mccormackDropDown
if salesPerson = Samantha Stikeleather
	gosub, mccormackDropDown
if salesPerson = Jeff Weller
	gosub, mccormackDropDown
if salesPerson = Jerry Holycross
	gosub, mccormackDropDown
if salesPerson = Dan Ciminelli
	gosub, mccormackDropDown
if salesPerson = Theresa Borio
	gosub, mccormackDropDown
if salesPerson = Cynthia (Cindy) Spittler
	gosub, mccormackDropDown
if salesPerson = Fred Simpson
	gosub, mccormackDropDown
if salesPerson = Gwyn Trojan
	gosub, mccormackDropDown
if salesPerson = Nick Hubbard (Nik)
	gosub, mccormackDropDown
if salesPerson = Kristen Luttner
	gosub, mccormackDropDown
if salesPerson = Patrick Bohman
	gosub, mccormackDropDown
if salesPerson = Eric Norviel
	gosub, mccormackDropDown
;========== END DOUG MCCORMACK =================

;=========== HEWITT ============
if salesPerson = Douglas Sears
	gosub, hewittDropDown
if salesPerson = Melissa Chandler
	gosub, hewittDropDown
if salesPerson = Hillary Tennant
	gosub, hewittDropDown
if salesPerson = Don Rathbauer
	gosub, hewittDropDown
if salesPerson = Tucker Lincoln
	gosub, hewittDropDown
if salesPerson = Joel Stradtner
	gosub, hewittDropDown
if salesPerson = Stephanie Koczur
	gosub, hewittDropDown
if salesPerson = Mike Hughes
	gosub, hewittDropDown
;=========== END HEWITT ============

;============ FOELS ================
if salesPerson = Larry Bellan
	gosub, foelsDropDown
if salesPerson = Kevin Clodfelter
	gosub, foelsDropDown
if salesPerson = Brian Thompson
	gosub, foelsDropDown
if salesPerson = Rashila Patel
	gosub, foelsDropDown
if salesPerson = Bill Balsanek
	gosub, foelsDropDown
if salesPerson = Chuck Costanza
	gosub, foelsDropDown
if salesPerson = Karl Kastner
	gosub, foelsDropDown
if salesPerson = Bob Riggs
	gosub, foelsDropDown
if salesPerson = Drew Smillie
	gosub, foelsDropDown
if salesPerson = Crystal Flowers
	gosub, foelsDropDown
;============ END FOELS ================

;========= SECOND ==============
if salesPerson = Helen Sun
	gosub, secondDropDown
if salesPerson = Steven Danielson
	gosub, secondDropDown
if salesPerson = Dominique Figueroa
	gosub, secondDropDown
if salesPerson = Jonathan McNally
	gosub, secondDropDown
if salesPerson = Yan Chen
	gosub, secondDropDown
if salesPerson = Katianna Pihakari
	gosub, secondDropDown
if salesPerson = Timothy Johnson
	gosub, secondDropDown
if salesPerson =  Dante Bencivengo 
	gosub, secondDropDown
if salesPerson =  Justin Chang 
	gosub, secondDropDown
;========= END SECOND ==============

;========= MCFADDEN ==============
if salesPerson = May Chou
	gosub, mcfaddenDropDown
if salesPerson = Steve Boyanoski
	gosub, mcfaddenDropDown
if salesPerson = Mark Woodworth
	gosub, mcfaddenDropDown
if salesPerson = Murray Fryman
	gosub, mcfaddenDropDown
if salesPerson = Lorraine Foglio
	gosub, mcfaddenDropDown
if salesPerson = Lauren Fischer
	gosub, mcfaddenDropDown
if salesPerson = Douglas McDowell
	gosub, mcfaddenDropDown
if salesPerson = Dana Stradtner
	gosub, mcfaddenDropDown
if salesPerson = Sarah Jackson
	gosub, mcfaddenDropDown
if salesPerson = Loris Fossir
	gosub, mcfaddenDropDown
if salesPerson = Mitch Lazaro
	gosub, mcfaddenDropDown
;========= END MCFADDEN ==============

;=========== BUTLER ==========
if salesPerson = Andrew Clark
	gosub, butlerDropDown
if salesPerson = Giovanni Pallante
	gosub, butlerDropDown
if salesPerson = David Kage
	gosub, butlerDropDown
if salesPerson = David Scott
	gosub, butlerDropDown
if salesPerson = Susan Gelman
	gosub, butlerDropDown
if salesPerson = Cari Randles
	gosub, butlerDropDown
if salesPerson = Sean Bennett
	gosub, butlerDropDown
if salesPerson = Daniel Quinn
	gosub, butlerDropDown
if salesPerson = Robin Sutka
	gosub, butlerDropDown
;=========== END BUTLER ==========

;=========== KLEIN ==========
if salesPerson = Susan Bird
	gosub, kleinDropDown
if salesPerson = Jerry Pappas
	gosub, kleinDropDown
if salesPerson = Jie Qian
	gosub, kleinDropDown
if salesPerson = Joe Bernholz
	gosub, kleinDropDown
if salesPerson = Yuriy Dunayevskiy
	gosub, kleinDropDown
if salesPerson = Nelson Huang
	gosub, kleinDropDown
if salesPerson = Kate Lincoln
	gosub, kleinDropDown
if salesPerson = Russ Constantineau
	gosub, kleinDropDown
;=========== END KLEIN ==========

;=========== CHEN ==========
if salesPerson = Haris Dzaferbegovic
	gosub, chenDropDown
if salesPerson = Donna Zwirner
	gosub, mccormackDropDown
;=========== END CHEN ==========

;=========== PORCH ==========
if salesPerson = Todd Stoner
	gosub, porchDropDown
if salesPerson = Jonathan Ferguson
	gosub, porchDropDown
if salesPerson = Nick Duczak
	gosub, porchDropDown
if salesPerson = Gerald Koncar
	gosub, porchDropDown
if salesPerson = Angelito Nepomuceno
	gosub, porchDropDown
if salesPerson = John Venesky
	gosub, porchDropDown
if salesPerson = Kristin Roberts
	gosub, porchDropDown
if salesPerson = David Kusel 
	gosub, porchDropDown
;=========== END PORCH ==========

;========== TOLLSTRUP ==========

if salesPerson = Taylor Graham  
	gosub, tollstrupDropDown
if salesPerson = Jerome Johemko  
	gosub, tollstrupDropDown
if salesPerson = Christie Baldizar 
	gosub, tollstrupDropDown

;========== END TOLLSTRUP ==========

;========== GAVINO ==========

if salesPerson = Christina Guintu  
	gosub, gavinoDropDown
if salesPerson = Ross Milam
	gosub, gavinoDropDown
if salesPerson = Sunny Chen  
	gosub, gavinoDropDown
if salesPerson = David Hill
	gosub, gavinoDropDown
if salesPerson = Elaine Miller
	gosub, gavinoDropDown
if salesPerson = Gary Scharrer
	gosub, gavinoDropDown
if salesPerson = Valerie Bruner
	gosub, gavinoDropDown
if salesPerson = Laura Howell
	gosub, gavinoDropDown
if salesPerson = Cecilia Snyder
	gosub, gavinoDropDown
if salesPerson = Tatiana Valle Melendez
	gosub, gavinoDropDown

;========== END GAVINO ==========

;======== CRAFTS =======

if salesPerson = Annie Cantelmo
	gosub, craftsDropDown
if salesPerson = Sitara Chauhan
	gosub, craftsDropDown
if salesPerson = Doug Meinhart
	gosub, craftsDropDown
if salesPerson = Tori Milioni
	gosub, craftsDropDown
if salesPerson = Bob Myers
	gosub, craftsDropDown
if salesPerson = Paulette Parker
	gosub, craftsDropDown
if salesPerson = Jane-Marie Kowalski
	gosub, craftsDropDown
if salesPerson = Ronsar Eid
	gosub, craftsDropDown
if salesPerson = Brian Ridley
	gosub, craftsDropDown 
if salesPerson = Krystina Simms
	gosub, craftsDropDown

;======== END CRAFTS =======


if salesPerson =
{
	GuiControl, Choose, salesManager, |1 
	GuiControl, Choose, managerCode, |1
	Gui Submit, NoHide
	GuiControl, Choose, salesDirector, |1
	GuiControl, Choose, directorCode, |1
	Gui Submit, NoHide
}
return