#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%
;The script compiler (ahk2exe) also supports library functions. 
;However, it requires that a copy of AutoHotkey.exe exist in the directory above the compiler directory (which is normally the case). If AutoHotkey.exe is absent, the compiler still works but library functions are not automatically included.

CoordMode Pixel, Screen
SendMode, Event
SetKeyDelay 100
openChangeSalesOrder()

InputBox salesConfirmation, SO#, Please enter a sales order number,,250,125,2376, 200 
if ErrorLevel
    Return

Send ^a{BackSpace}
Clipboard := salesConfirmation
Send % Clipboard
Send !s
Sleep 200
Send u
WinWaitActive Output Details
Send {down 2}+{Space}^{Tab}{Tab 3}{enter}
WinWaitActive Print:
Send ^+{tab}{tab}{Enter}
WinWaitActive ,,Custom  Container
Sleep 1000
Send ^{tab 3}
Sleep 500
Send {Esc}
Sleep 500
Send ^{Tab}{Enter}
WinWaitActive Save As
WinWaitActive ,,Custom  Container
Send {F3}
WinWaitActive Output Details
Send {Esc}
Return



#Include OrderOrganizerFunctions.ahk