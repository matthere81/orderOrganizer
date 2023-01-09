#NoEnv
#SingleInstance, Force
SendMode, Input
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%


;**** Opens the Change Sales Order in SAP (VA02) ****
openChangeSalesOrder() {
    SetTitleMatchMode, 3
    IfWinExist Change Sales Order: Initial Screen
        {
            WinActivate Change Sales Order: Initial Screen
        } 
    else IfWinNotExist Change Sales Order: Initial Screen
        {
            WinActivate SAP Easy Access
            WinWaitActive SAP Easy Access
            ControlFocus Edit1, SAP Easy Access
            Clipboard := "VA02"
            Send %Clipboard%{Enter}
        }
}
Return