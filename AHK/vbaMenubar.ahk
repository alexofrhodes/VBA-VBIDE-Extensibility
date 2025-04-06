/* Info
AUTHOR: 
    Anastasiou Alex
    anastasioualex@gmail.com
    https://github.com/alexofrhodes <-- Repos
    https://alexofrhodes.github.io <-- Blog
    https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg

PURPOSE: GUI to run Excel macros with a menu bar
*/

#SingleInstance, force
SetWorkingDir %A_ScriptDir%
#include LIB\ini-editor.ahk

#Persistent

#Include LIB\Class Dock.ahk

callbackRun := false
snippetsID := 0

LoadOptions:
; IniRead, WorkbookName, config.ini, Settings, WorkbookName    ;moved to runExcelMacro
IniRead, MenuFile, config.ini, Settings, MenuFile
IniRead, ItemsPerColumn, config.ini, Settings, ItemsPerColumn
IniRead, fontSize, config.ini, Settings, fontSize

; Gui,Font, s%fontSize%, Arial

IniRead, xPos, config.ini, Settings, xPos, 100
IniRead, yPos, config.ini, Settings, yPos, 100

IniRead, myHotkey, config.ini, Settings, myHotkey
Hotkey, %myHotkey%,Start

ItemsPerColumn := ItemsPerColumn + 0

;Custom Tray Icon
I_Icon = %A_WorkingDir%\vbaGUI.ico ;dAKirby309 (Michael) at https://icon-icons.com/icon/excel-mac/23559
IfExist, %I_Icon%
	Menu, Tray, Icon, %I_Icon%

;Event for Tray icon left click
OnMessage(0x404, "AHK_NOTIFYICON")
AHK_NOTIFYICON(wParam, lParam)
{
    if (lParam = 0x201) ; WM_LBUTTONDOWN
    {
        Gui, Show, AutoSize w600 x%xpos% y%ypos% , VBA Callbacks
        return 0
    }
}

; Main
start:

gui, destroy
Gui, -Caption +ToolWindow +0x400000
Gui, +AlwaysOnTop 
menuData := GetMenuData()
if !menuData {
    MsgBox, No valid menu data found.
    ExitApp
}


CreateMenuBar(menuData)
Return

; Read menu file and parse tabs and content
GetMenuData() {
    global
    FileRead, content, %MenuFile%
    if ErrorLevel {
        MsgBox, Error reading menu file.
        ExitApp
    }

    menuData := {}
    currentTab := ""

    ; Parse the file line by line
    for each, line in StrSplit(content, "`n", "`r") {
        if (line = "")
            continue ; Skip blank lines

        FirstCharacter := SubStr(line, 1, 1)
        if FirstCharacter in !,-,.,;
            continue

        ; If line starts without indentation, treat it as a new menu item group
        if not (FirstCharacter = A_Tab) {
            currentTab := line
            menuData[currentTab] := []
        } else if currentTab {
            menuData[currentTab].Push(Trim(SubStr(line, 2)))
        }
    }
    return menuData
}

; Create GUI with dynamically loaded menu bar
CreateMenuBar(menuData) {
    global

    
    ; Add dynamic menus for each tab
    for tabName, items in menuData {
        Menu, %tabName%, Add  ; Create a submenu for each tab, creates a horizontal separator, dunno why we need to add this for the column divider to work
        
        itemCount := 0
        for each, macroName in items {
            if macroName = ""
                continue
            ; Add each macro name as a menu item
            Menu, %tabName%, Add, %macroName%, RunExcelMacro

            itemCount++

            ; Check if we need to start a new column
          
            if (itemCount > ItemsPerColumn) {
                
                MenuHandle := GetMenuHandle( tabName ) 
    
                VarSetCapacity(mii, cb:=16+8*A_PtrSize, 0) ; A_PtrSize is used for 64-bit compatibility.
                NumPut(cb, mii, "uint")
                NumPut(0x100, mii, 4, "uint") ; fMask = MIIM_FTYPE
                NumPut(0x20, mii, 8, "uint") ; fType = MFT_MENUBARBREAK
                DllCall("SetMenuItemInfo", "ptr", MenuHandle, "uint", itemCount, "int", 1, "ptr", &mii)
    
                itemCount := 0  ; Reset the counter for the new column
            }            
        }
        Menu, %tabName%, delete,1&  ; need to delete that first horizontal separator.. 

        Menu, MyMenuBar, Add, %tabName%, :%tabName% ; Attach the submenu to the main menu
    }

    Menu, MyMenuBar, Add,  SNIPPETS, Snippets

    Menu, guiControls, Add,  Options, MenuSettings
    Menu, guiControls, Add,  Reload, ReloadMe
    menu, guicontrols, add, Hide, GuiClose

    menu, mymenubar, add, guiControls, :guiControls


    ; Attach menu to GUI
    Gui, Menu, MyMenuBar
   
    ; Gui, Add, Button, x10 w500 gExit, Exit This Example

    Gui, Add, Text, xs y5 w800 h1 0x7  ;Horizontal Line > Black

    Gui, +hwndGuihwnd
    Gui, Show, AutoSize x%xpos% y%ypos% , VBA Callbacks
    gui, hide
    SetTimer, CheckVBEActive, 500  ; Check every 500ms

    OnMessage(0x0201, "WM_LBUTTONDOWN")
    Return
}

Snippets:
    Run, % A_ScriptDir "\snippets\snippets.exe",,, snippetsID
    ; TrayTip , Snippets is running, Press 2 x CTRL to show `n right click on tray for more info  ;, Timeout, Options
Return

CheckVBEActive:
    WinGet, activeWindow, ID, A  ; Get the ID of the active window
    WinGetClass, className, ahk_id %activeWindow%

    ; Check if the active window is the VBA editor
    if (className = "wndclass_desked_gsk") {
        if (!callbackRun) {  
            ; Run the callback only if it hasn't been executed yet
            ; Set flag to true to prevent re-running
            callbackRun := true

            WinGet, hwnd, ID, ahk_class wndclass_desked_gsk
            ;The first argument is the host's hwnd and the second the client's hwnd
            Gui, Show, AutoSize x%xpos% y%ypos% , VBA Callbacks
            SetTimer, CheckVBEActive, off
            exDock := new Dock( hwnd,Guihwnd)
            exDock.Position("BL")
            exDock.CloseCallback := Func("CloseCallback")
            ; gosub, snippets
        }
    }

Return


CloseCallback(self)
{
	WinKill, % "ahk_id " self.hwnd.Client
    Process, Close, % snippetsID
	ExitApp
}



WM_LBUTTONDOWN() {
	PostMessage, 0xA1, 2,,, A
}

GetMenuHandle(menu_name) ;from MenuIcons v2
{
    static   h_menuDummy
    ; v2.2: Check for !h_menuDummy instead of h_menuDummy="" in case init failed last time.
    If !h_menuDummy
    {
        Menu, menuDummy, Add
        Menu, menuDummy, DeleteAll

        Gui, 99:Menu, menuDummy
        ; v2.2: Use LastFound method instead of window title. [Thanks animeaime.]
        Gui, 99:+LastFound

        h_menuDummy := DllCall("GetMenu", "uint", WinExist())

        Gui, 99:Menu
        Gui, 99:Destroy
        
        ; v2.2: Return only after cleaning up. [Thanks animeaime.]
        if !h_menuDummy
            return 0
    }

    Menu, menuDummy, Add, :%menu_name%
    h_menu := DllCall( "GetSubMenu", "uint", h_menuDummy, "int", 0 )
    DllCall( "RemoveMenu", "uint", h_menuDummy, "uint", 0, "uint", 0x400 )
    Menu, menuDummy, Delete, :%menu_name%
    
    return h_menu
}

;Close GUI when exit button pressed or ESC pressed. This doesn't stop the script's execution.
; GuiEscape:
GuiClose:
	gosub SavePos
	Gui, hide

    callbackRun := false
    SetTimer, CheckVBEActive, 500  ; Check every 500ms
return

SavePos:
    Gui +lastfound
    WinGetPos, xPos, yPos
	if (xPos <= 0) 
		return
    IniWrite, %xPos%, config.ini, Settings, xPos
    IniWrite, %yPos%, config.ini, Settings, yPos
Return

MenuSettings:
	IniSettingsEditor("Settings", "config.ini")
Return

ReloadMe:
	gosub SavePos
	Reload
	Sleep 1000 ; If successful, the reload will close this instance during the Sleep, so the line below will never be reached.
	MsgBox, 4,, The script could not be reloaded. Would you like to open it for editing?
	IfMsgBox, Yes, Edit
return

#Include, LIB\excel_get.ahk
#include, LIB\runeXcelMacro.ahk

