#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

;**** Opens the Change Sales Order in SAP (VA02) ****
; openChangeSalesOrder() {
;     SetTitleMatchMode, 3
;     IfWinExist Change Sales Order: Initial Screen
;         {
;             WinActivate Change Sales Order: Initial Screen
;         } 
;     else IfWinNotExist Change Sales Order: Initial Screen
;         {
;             WinActivate SAP Easy Access
;             WinWaitActive SAP Easy Access
;             ControlFocus Edit1, SAP Easy Access
;             Clipboard := "VA02"
;             Send %Clipboard%{Enter}
;         }
; }
; Return

forwardSoftwareLicense()
{
    outlookApp := ComObjActive("Outlook.Application") ; Get running instance of Outlook
    namespace := outlookApp.GetNamespace("MAPI") ; Get MAPI namespace
    
    folder := namespace.GetDefaultFolder(6).Folders("SW Licenses") ; 6 corresponds to the Inbox
    
    unreadEmails := folder.Items.Restrict("[Unread] = true") ; Get all unread emails
    
    for email in unreadEmails {
        body := email.Body ; Get the body of the email
        subject := email.Subject ; Get the subject of the email
        RegExMatch(body, "SO/PO:\s*([a-zA-Z0-9]+)", match) ; Find the alphanumeric string after "SO/PO:"
        ; MsgBox, % "Subject: " subject "`nFound string: " match1 ; Display the subject and the string
    }
}

forwardSoftwareLicense()