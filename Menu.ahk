
; Create the submenu
Menu FileMenu, Add, &Open, readtheini
Menu ViewMenu, Add, Show Checklists, ShowChecklists
; Menu, SettingsMenu, Add, Show Checklists, ShowChecklists

; Add the submenu to the main menu
Menu MainMenu, Add, &File, :FileMenu
; Menu MainMenu, Add, &Settings, :SettingsMenu
Menu MainMenu, Add, &View, :ViewMenu


Gui Menu, MainMenu



; From Documentation
; Menu, FileMenu, Add, Script Icon, MenuHandler
; Menu, FileMenu, Add, Suspend Icon, MenuHandler
; Menu, FileMenu, Add, Pause Icon, MenuHandler
; Menu, MyMenuBar, Add, &File, :FileMenu
; Gui, Menu, MyMenuBar
; ; Gui, Add, Button, gExit, Exit This Example
; Gui, Show
; return

; MenuHandler:
; return