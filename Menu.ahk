
; Create the submenu
Menu SubMenu, Add, &Open, readtheini
; Menu, SubMenu, Add, SubItem2, SubItem2Handler

; Add the submenu to the main menu
Menu MainMenu, Add, &File, :SubMenu

Gui Menu, MainMenu

