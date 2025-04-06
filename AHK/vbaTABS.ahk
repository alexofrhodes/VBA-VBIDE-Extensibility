/*	Info
AUTHOR: 
	Anastasiou Alex
		anastasioualex@gmail.com
		https://github.com/alexofrhodes <-- Repos
		https://alexofrhodes.github.io	<-- Blog
		https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
		
PURPOSE: GUI to run excel macros
*/


;#IfWinActive, ahk_exe EXCEL.EXE			;if any excel window is active
;#IfWinActive ahk_class wndclass_desked_gsk ;if vbeditor window is active

#SingleInstance, force
SetWorkingDir %A_ScriptDir%
#include LIB\ini-editor.ahk
#Include LIB\ImageButton.ahk
EStyle := [[0, 0x80F0F0F0, , , 8, 0xFFF0F0F0, 0x8046B8DA, 2] ; normal
		, [0, 0x80C6E9F4, , , 8, 0xFFF0F0F0, 0x8046B8DA, 2] ; hover
		, [0, 0x8086D0E7, , , 8, 0xFFF0F0F0, 0x8046B8DA, 2] ; pressed
		, [0, 0x80F0F0F0, , , 8, 0xFFF0F0F0, 0x8046B8DA, 2]]




LoadOptions:
; IniRead, WorkbookName, config.ini, Settings, WorkbookName    ;moved to runExcelMacro
IniRead, MenuFile, config.ini, Settings, MenuFile
IniRead, ItemsPerColumn, config.ini, Settings, ItemsPerColumn
IniRead, fontSize, config.ini, Settings, fontSize
IniRead, xPos, config.ini, Settings, xPos, 100
IniRead, yPos, config.ini, Settings, yPos, 100

IniRead, myHotkey, config.ini, Settings, myHotkey
Hotkey, %myHotkey%,Start


ItemsPerColumn := ItemsPerColumn + 0
counter := 0


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
        gui,show
        return 0
    }
}
    
gui, destroy
Gui, +AlwaysOnTop 
menuData := GetMenuData()
if !menuData {
    MsgBox, No valid menu data found.
    ExitApp
}

Gui, Font, s%fontSize%, Arial
gui, add, button, gMenuSettings,  Options
gui, add, button, ys gReloadMe,Reload

CreateGuiWithTabs(menuData)
Gui, Tab ; End Tab section
Return

Start:
    Gui, Show, x%xpos% y%ypos% AutoSize, vba Callbacks
return

; @TODO resize form 
tabChange:
    beforeTab:=CurrentTab
    Gui,Submit,NoHide
    afterTab:=CurrentTab
    ; MsgBox % A_GuiEvent "." A_EventInfo "`nbeforeTab: " beforeTab " : after: " afterTab
    ; If (beforeTab=1 && beforeTab != afterTab)
        ; MsgBox Jumped from %beforeTab% to %afterTab%
Return


; Step 1: Read menu file and parse tabs and content
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
        ; line := Trim(line)
        if (line = "")
            continue ; Skip blank lines

        FirstCharacter := SubStr(line, 1, 1)
		if firstCharacter in !,-,.,;
			Continue

        ; If line starts without indentation, treat it as a new tab
        if not (FirstCharacter = A_Tab) 
        {
            currentTab := line
            menuData[currentTab] := []
        } else if currentTab {
            menuData[currentTab].Push(Trim(SubStr(line, 2)))
        }
    }
    return menuData
}

; Step 2: Create GUI with dynamically loaded tabs and buttons
CreateGuiWithTabs(menuData) {
    global
    

    ; Create tabs dynamically
    tabNames := ""
    for tabName in menuData
        tabNames .= (tabNames = "" ? "" : "|") . tabName

    Gui, Add, Tab3 ,  Buttons xs gtabChange vCurrentTab AltSubmit, %tabNames%

    ; Add buttons to each tab
    for tabName in menuData {
        Gui, Tab, %tabName%
        x := 20, y := 120, counter := 0
        MenuFound:=0
        for each, line in StrSplit(content, "`n", "`r")
        {
            if !line and (menuFound = 0) ; Check if line is blank or unset
                Continue

            if (line=tabName)
            {
                MenuFound++
                continue
            }	
            
            if !menufound
                Continue
            FirstCharacter:= substr(line,1,1)
            if firstCharacter in !,-,.,;
                continue
            
            if not (firstCharacter=A_Tab) && (trim(line)<>"")
                Break

            Position:= InStr(line, ";")
            if Position>0
                line:=SubStr(line, 1, Position - 1)
            Position:= InStr(line, "!")		;because ezMenu uses ! key to switch the GoSub
            if Position>0
                line:=SubStr(line, 1, Position - 1)
            line:=trim(line)
            
            ;alternatively to force new column leave empty line (resets counter for the rest of the controls in the new column)
            if !line 
            {
                counter := 0
                x += 180, y := 120
                continue
            } 


            if (counter = ItemsPerColumn)
            {
                counter := 0
                x += 180, y := 120
            } else if (counter >0) {
                y += 30
            }
    
              counter++
          Gui, Add, Button, w120 x%x% y%y% gRunExcelMacro hwndBtn, %line% ; gRunExcelMacro
            
            ; ImageButton.Create(Btn, line, EStyle*) 
        }
    }

    Return
}



ReloadMe:
	gosub SavePos
	Reload
	Sleep 1000 ; If successful, the reload will close this instance during the Sleep, so the line below will never be reached.
	MsgBox, 4,, The script could not be reloaded. Would you like to open it for editing?
	IfMsgBox, Yes, Edit
return

MenuSettings:
	IniSettingsEditor("Settings", "config.ini")
Return

EditFile:
	run %MenuFile%
return

;Close GUI when exit button pressed or ESC pressed. This doesn't stop the script's execution.
GuiEscape:
GuiClose:
	gosub SavePos
	Gui, hide
return

SavePos:
    Gui +lastfound
    WinGetPos, xPos, yPos
	if (xPos <= 0) 
		return
    IniWrite, %xPos%, config.ini, Settings, xPos
    IniWrite, %yPos%, config.ini, Settings, yPos
Return



#Include, LIB\runeXcelMacro.ahk
#include, LIB\excel_get.ahk