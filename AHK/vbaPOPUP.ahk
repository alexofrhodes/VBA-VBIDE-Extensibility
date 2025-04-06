#noEnv
#singleinstance, force
sendMode input
setWorkingDir, % a_scriptDir

;#IfWinActive ahk_exe EXCEL.EXE ;sometimes error
;#IfWinActive ahk_class XLMAIN ;sometimes error\

;#IfWinActive ahk_class wndclass_desked_gsk ;if vbeditor window is active

/* NOTES
	----------
	Replace WorkbookName:="ProjectStarter.xlam!" with your workbook where the macros are.
	Edit/create the vba.menu file in same folder as this file to contain your macros. Levels by tab indentation.
	ezMenu will create a menu from the vba.menu file
	info on ezMenu at https://github.com/davebrny/ezMenu
*/
#include, LIB\ini-editor.ahk

I_Icon = %A_WorkingDir%\vbaGUI.ico ;dAKirby309 (Michael) at https://icon-icons.com/icon/excel-mac/23559
IfExist, %I_Icon%
	Menu, Tray, Icon, %I_Icon%


LoadOptions:
	IniRead, myHotkey, config.ini, Settings, myHotkey
	IniRead, WorkbookName, config.ini, Settings, WorkbookName
	IniRead, MenuFile, config.ini, Settings, MenuFile
	Hotkey, %myHotkey%,Start
Return

Start: 
    ezMenu("vbaMenu", MenuFile)
    ;Reload ;otherwise throws error (menu item has no parent)
return

MenuSettings:
	IniSettingsEditor("Settings", "config.ini")
Return

#Include, LIB\Excel_get.ahk
#Include, Lib\RunExcelMacro.ahk

#Include, LIB\ezmenu.ahk
