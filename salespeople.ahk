#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

; Path to the Excel workbook
workbookPath := "C:\Users\" . A_UserName . "\OneDrive - Thermo Fisher Scientific\General\Training Docs\Sales List fed 12.xlsx"
salesWorkBookName := "Sales List fed 12.xlsx"

; Function to open an Excel workbook
OpenExcelWorkbook(path) {
    xlApp := ComObjCreate("Excel.Application")
    xlApp.Visible := false
    workbook := xlApp.Workbooks.Open(path)
    return workbook
}

wb := OpenExcelWorkbook(workbookPath)

winOpen := 0

if WinExist(salesWorkBookName) {
    WinClose
    winOpen := 1
}

;-------- LIST OUT WORKSHEET NAMES --------;
ws1 := wb.Worksheets[1] ; - West Denise Schwartz
ws2 := wb.Worksheets[2] ; - East & Canada Maroun
ws3 := wb.Worksheets[3] ; - Digital
ws4 := wb.Worksheets[4] ; - IOMS Sales
;-------- END LIST OUT WORKSHEET NAMES END --------;

; Identify sales regions
ws1Regions := ["AMER-GCM", "AMER-MWM", "AMER-MOM", "AMER-NWM", "AMER-SWM"]
ws2Regions := ["AMER-CAM", "AMER-TSM", "AMER-NEM", "AMER-MAM"]
ws3Regions := ["USA"]
ws4Regions := ["IOMS Sales Team"]
salesManagers := []
salesDirectors := []
salesPeople := []
regionsArray := [ws1Regions, ws2Regions, ws3Regions, ws4Regions]
worksheetsArray := [ws1, ws2, ws3, ws4] ; ws1, ws2, 

; Get the last row and column in the worksheet
; lastRow := ws1.UsedRange.Rows.Count
; lastColumn := ws1.UsedRange.Columns.Count

; Arrays for storing names
names_with_white_bg := []
names_with_non_white_bg := []

for index, worksheet in worksheetsArray
{
    ; cellA3 := worksheet.Range("A3")
    ; MsgBox, % "Value of cell A3: " . cellA3.Value
    ; MsgBox, % "Formula of cell A3: " . cellA3.Formula
    ; MsgBox, % "Address of cell A3: " . cellA3.Address
    ; MsgBox, % "Color of cell A3: " . cellA3.Interior.Color
	; MsgBox, % "The used range of " . worksheet.Name . " is " . usedRange.Address
	usedRange := worksheet.UsedRange
	
	for column in usedRange.Columns
	{
    	for cell in column.Cells
    	{		
			; MsgBox, 1,, % cell.Value . " " . found . " at " cell.Address . " " . cell.Interior.Color ;. " " . foundDirector
			; IfMsgBox, Cancel
			; {
			; 	break	
			; }	
			; Find names
			found := RegExMatch(cell.Value, "(?:[a-zA-z]*\s([\(][a-zA-z]*[\)]\s[a-zA-Z]*|[a-zA-Z]*))")
			; found := RegExMatch(cell.Value, "(?:[B{1}][a-zA-z]*\s[a-zA-z]*)")
			cell.Value := Trim(cell.Value)
			if (cell.Value = "Sales Director:")
				{
					myOffset := cell.Offset(1,0)
					foundDirector := myOffset.Value
					salesDirectors.Push(foundDirector)
				}
			; if (found = 1) && !(cell.Value = " ") && !(cell.Interior.Color = 16777215) && !(cell.Value = "Gulf Coast") && !(cell.Value = "Sales Director:")
			; 	&& !(cell.Value = "Manager TBD") && !(cell.Value = "Sales Director:") && !(cell.Value = foundDirector)
			; 	MsgBox, 1,, % cell.Value . " " . found . " at " cell.Address . " " . cell.Interior.Color . " " . foundDirector
			; 	IfMsgBox, Cancel
			; 	{
			; 		break	
			; 	}	
    	}
	}		
}


; Print the values in the endingCells array
for index, value in salesDirectors
    display .= value . "`n"

MsgBox % display

if winOpen = 1
{
    ; Open workbook
    wb := xl.Workbooks.Open(workbookPath)
    xl.Visible := True
}
else if winOpen = 0
{
    wb.Close()
    xl.quit()
}

; Perform any desired operations on the names arrays
; MsgBox("Names with white background:")
; for index, name in names_with_white_bg {
;     MsgBox % "Names with white backgroudnd " . name
; }

; ; MsgBox("Names with non-white background:")
; for index, name in names_with_non_white_bg {
;     MsgBox % "Names with non-white background: " . name
; }

MsgBox, end
Return


;-------- FUNCTIONS FUNCTIONS FUNCTIONS --------;

; OpenExcelWorkbook(workbookPath)
; {
;     xl := ComObjCreate("Excel.Application")
;     xl.Visible := False
;     wb := xl.Workbooks.Open(workbookPath)
;     return wb
; }

; Function to find values in an array in an Excel worksheet
FindArrayValues(array, worksheet)
{
    ; Loop through the array
    for i, value in array
    {
        ; Search for the value in the worksheet
        cell := worksheet.Cells.Find(value)
        cellAddress := cell.Address
		; Get the interior color of the cell
		interiorColor := cell.Interior.color

		; Search for the next cell with the same interior color in the same column
        nextCell := cell.Offset(1, 0)
		nextCellColor := nextCell.Interior.color

		; Iterate over each cell in the same column until there is a cell with no value found
        while (nextCell.Value != "")
        {
			msgbox % nextCell.Value
			nextCell := nextCell.Offset(1, 0)
            ; if (nextCellColor = interiorColor)
            ; {
            ;     if (nextCellColor != -4142)
            ;     {
            ;         MsgBox, % "Value: " . nextCell.Value . " from first else`nNextCellColor is " . nextCellColor . " and original color is " . interiorColor
            ;     }
            ;     nextCell := nextCell.Offset(1, 0)
            ;     nextCellColor := nextCell.Interior.color
            ; }
            ; else
            ; {
            ;     if (nextCellColor != -4142)
            ;     {
            ;         MsgBox, % "Value: " . nextCell.Value . "`nNextCellColor is " . nextCellColor . " and original color is " . interiorColor
            ;     }
            ;     break
            ; }
		}
    }
}



LoopRegions(ws1Regions)
{
	; Initialize array to store cell information dictionaries
    cellInfoList := []
    ; Loop through each value in the array
    for eachRegion in ws1Regions
    {
        foundCell := ws1.Cells.Find(ws1Regions[eachRegion])
		MsgBox, Found cell: %foundCell%
        currentRegion := ws1Regions[eachRegion]
		cellInfo := GetCellInfo(foundCell)
		cellInfoList.push(cellInfo)
    }

	; Loop over cell information dictionaries and show address in message box
    for eachCell in cellInfoList
    {
        MsgBox, % eachCell["Address"]
    }

    ; }

        
	; Add the cellInfo dictionary to the colorRanges array
	colorRanges.Push(cellInfo)

	; Loop through each value in the array
	; for eachRegion in ws1Regions
	; {
	;     foundCell := ws1.Cells.Find(ws1Regions[eachRegion])
	;     currentRegion := ws1Regions[eachRegion]
	; 	MsgBox, % ws1Regions[eachRegion]
	; }  
}

GetCellInfo(foundCell)
{
    foundAddress := foundCell.Address
    foundColor := foundCell.Interior.Color
    nextCell := foundCell.Offset(1, 0)
    
    ; Return a dictionary containing the cell information
    cellInfo := {}
    cellInfo["Address"] := foundAddress
    cellInfo["Color"] := foundColor
    cellInfo["NextCell"] := nextCell
	MsgBox, Cell info: %cellInfo%
    
    return cellInfo

	    ; if (foundCell <> "")
	; {
	; 	; Get the address of the found cell
	; 	foundAddress := foundCell.Address
		
	; 	; Get the interior color of the found cell
	; 	foundColor := foundCell.Interior.Color

	; 	; Get the next cell in the same row
	; 	nextCell := foundCell.Offset(1, 0)
}

/*
* Given a starting cell and an interior color, finds the next cell with the same interior color in the same column
* Returns the next cell and its interior color, or an empty string and 0 if not found
*/
FindNextCellWithSameInteriorColor(startingCell, interiorColor, worksheet)
{
    foundCell := startingCell.Offset(1, 0)
    foundColor := foundCell.Interior.Color
    
    while (foundCell.Column = startingCell.Column)
    {
        if (foundColor = interiorColor)
        {
            return [foundCell, foundColor]
        }
        else
        {
            foundCell := foundCell.Offset(1, 0)
            foundColor := foundCell.Interior.Color
        }
    }
    
    ; If the loop completes without finding a match, return empty string and 0
    return ["", 0]


	; 	; Loop through the cells in the same row until you find one with the same interior color
	; 	while (nextCell.Column = foundCell.Column)
	; 	{
	; 		; Get the interior color of the next cell
	; 		nextColor := nextCell.Interior.Color

	; 		if (nextColor = foundColor)
	; 		{
}

/*
* Given a starting cell, starting row, ending row, and worksheet, extracts sales people from a range of cells
* Returns an array of sales people
*/
ExtractSalesPeopleInRange(startingCell, startingRow, endingRow, worksheet)
{
    salesPeople := []
    
    Loop % (endingRow - startingRow + 1)
    {
        cellValue := worksheet.Cells(startingRow + A_Index - 1, startingCell.Column).Value
        cellValue := Trim(cellValue)
        
        if (cellValue = "TBD")
            break
        
        salesPeople.Push(cellValue)
    }
    
    return salesPeople
}


/*
* Given a starting cell and worksheet, finds the next cell with the same interior color in the same column,
* and extracts sales people from the range of cells between the starting and ending cells
* Returns an array of sales people and the sales manager, or an empty array if the manager value is "Manager TBD"
*/
; FindSalesPeopleInRange(startingCell, worksheet)
; {
;     salesPeople := []
;     foundColor := startingCell.Interior.Color
    
;     while (startingCell <> "")
;     {
;         [nextCell, nextColor] := FindNextCellWithSameInteriorColor(startingCell, foundColor, worksheet)
        
;         if (nextCell = "")
;             break
        
;         startingRow := startingCell.Row + 1
;         endingRow := nextCell.Row - 1
        
;         extractedSalesPeople := ExtractSalesPeopleInRange(startingCell, startingRow, endingRow, worksheet)
;         salesPeople.Push(extractedSalesPeople)
        
;         managerValue := nextCell.Value
;         managerValue := Trim(managerValue)
        
;         if (managerValue = "Manager TBD")
;             break
        
;         salesManagers.Push(managerValue)
        
;         startingCell := nextCell
;         foundColor := nextColor
;     }
    
;     return [salesPeople, salesManagers]
; }

;-------- END FUNCTIONS END FUNCTIONS END FUNCTIONS --------;


; Display all the values in the array
; for Index, Value in Data
; 	Display .= Index "`t" Value ;"`n"
; MsgBox % Display

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