#SingleInstance Force

global MyMenu
global BaseFolder := A_ScriptDir . "\Snippets"
global extensions := ["txt"]

#Include lib\FolderStructure.ahk
#Include lib\customTray.ahk

Tray := A_TrayMenu
Tray.Add()  ; Creates a separator line.
for k,v in ["CTRL x2`tShow Menu",
			"Click File`tPaste Content",
			"CTRL Click`tEdit File",
			"SHIFT Click`tAppend to file"]
			{
				Tray.Add(v, DoNothing)
				Tray.SetIcon(v, "icons\hotkey.ico")
				tray.Disable(v)
			}
Tray.Add()
SetupTray()

;-----------------------
;Listen to hotkeys if vbeditor is active window
;-----------------------

	; #HotIf WinActive("ahk_class wndclass_desked_gsk")
	
;-----------------------
;Long press right button
;-----------------------

	; RButton:: 
	; {
	; 	startTime := A_TickCount 
	; 	KeyWait("RButton", "U")  
	; 	keypressDuration := A_TickCount-startTime 
	; 	if (keypressDuration > 200) 
	; 	{
	; 		Main()
	; 	}
	; 	else 
	; 	{
	; 		Send("{RButton}")
	; 	}        
	; }

;-----------------------	
;Double press ctrl
;-----------------------

~Ctrl Up:: 
{
	If A_ThisHotkey = A_PriorHotkey && A_TimeSincePriorHotkey < 400
		Main()
}

Main(){
	global
	try	
		myMenu.Delete
	myMenu:= Menu()	
	AddFolderStructureToMenu(MyMenu, BaseFolder, extensions, "theHandlerFunction")																				
	myMenu.Show()
}



; | Modifier Key   | Action   |
; |---------------|-----------|
; | Ctrl + Shift  | Overwrite |
; | Ctrl          | Edit      |
; | Shift         | Append    |
; | (None)        | Paste     |


theHandlerFunction(filePath, *) {
    if GetKeyState("Ctrl") && GetKeyState("Shift") { 
        ; Overwrite the file with clipboard content
        Send("{Ctrl down}c{Ctrl up}")
        Sleep(100)
        text := "`n" A_Clipboard        
		FileDelete(filePath)  ; Delete the file before overwriting
        FileAppend(A_Clipboard, filePath)
    } else if GetKeyState("Ctrl") { 
        ; Open the file in an editor
        Run 'edit ' filePath
    } else if GetKeyState("Shift") { 
        ; Append clipboard content to the file
        Send("{Ctrl down}c{Ctrl up}")
        Sleep(100)
        text := "`n" A_Clipboard
        FileAppend(text, filePath)
    } else { 
        ; Copy the file content and paste it
        text := FileRead(filePath)
        A_Clipboard := text
        Sleep(100)
        Send("{Ctrl down}v{Ctrl up}")
    }
}




