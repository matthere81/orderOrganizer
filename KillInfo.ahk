#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
#Persistent
#SingleInstance Force

; Click through info windows

SetTimer, StartOver, 60000

Start:
   SendMode, Event
   Setkeydelay, 250
   SetTitleMatchMode, 2
   GroupAdd, AutoEnter, Information,,,Restrict Value Range (1)
   ; GroupAdd, AutoEnter, Information,,,Message
   ; GroupAdd, GroupName, WinTitle [, WinText, Label, ExcludeTitle, ExcludeText]
   ; GroupAdd, AutoEnter, Information,,,License Information Multiple Logon,
   GroupAdd, AutoEnter, Internal Text Only,,,Restrict Value Range (1)
   GroupAdd, AutoEnter, Service Offering - HLL required,,,Restrict Value Range (1)
   GroupAdd, AutoEnter, Service Offering - Material Validation,,,Restrict Value Range (1)
   GroupAdd, AutoEnter, Service Offering - Quantity,,,Restrict Value Range (1)
   GroupAdd, AutoEnter, Service Offering - Serial Tracking,,,Restrict Value Range (1)
   GroupAdd, AutoEnter, Contract Plan Material - CRM,,,Restrict Value Range (1)
   WinWaitActive, ahk_group AutoEnter,
   ; IfWinActive, ahk_group AutoEnter,
   IfWinExist, ahk_group AutoEnter,
      loop,
   {
      Sleep, 100
      IfWinActive, Service Offering - HLL required,,,Restrict Value Range (1)
      {
         Send {Tab}{Enter}
      }
      ControlSend, Button2, {ENTER}, ahk_group AutoEnter,
      Sleep, 250
      IfWinNotExist, ahk_group AutoEnter,
      {
         gosub, Start
      }
      
   }

return

StartOver:
   Reload
return

F11::Pause

::ncht::ITEMS SENT AT NO CHARGE PER *SALES REP* FOR *CONTACT* / PHONE: 

   ;~ Start:
   ;~ SetTitleMatchMode, 2
   ;~ IfWinActive, information
   ;~ WinWaitActive, information
   ;~ MsgBox, here
   ;~ loop,
   ;~ {	
   ;~ Sleep, 100
   ;~ ControlSend,, a, information
   ;~ Sleep, 500
   ;~ IfWinNotActive, ahk_group AutoEnter,
   ;~ gosub, Start
   ;~ }
   ;~ return

