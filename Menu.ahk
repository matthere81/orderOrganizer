
; Create the submenu

; ---- File Menu ----
Menu FileMenu, Add, &Open, OpenFileFromMenu
Menu FileMenu, Add, &Save, SaveToIni
Menu FileMenu, Add, &Delete File, DeleteFile
; ---- End File Menu ----

; ---- View Menu ----
Menu ViewMenu, Add, &Show Checklists, ShowChecklists
; ---- End View Menu ----

; ---- Settings Menu ----
Menu SettingsMenu, Add, &Archive Files, ArchiveFiles
; ---- End Settings Menu ----

; Add the submenu to the main menu
Menu MainMenu, Add, &File, :FileMenu
Menu MainMenu, Add, &View, :ViewMenu
Menu MainMenu, Add, &Settings, :SettingsMenu

Gui Menu, MainMenu
