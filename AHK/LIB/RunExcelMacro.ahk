
RunExcelMacro:
{
	try {
		XL := Excel_Get()
	} catch {
		MsgBox, 16,, Can't obtain Excel! 
		return
	}
	;MsgBox, 64,, Excel obtained successfully!   ;for debugging purposes
	
	IniRead, WorkbookName, config.ini, Settings, WorkbookName

	;a space is allowed in the following format: [&1] MacroName
	;to allow a GUI accelerator between the braces eg. [accelerator]
	if instr(A_GuiControl, A_Space){
		StringSplit, Procedure, A_GuiControl, %A_Space%
		macro:= "'" . WorkbookName . "'" . "!" . Procedure2 
	}else{
		macro:= "'" . WorkbookName . "'" . "!" . A_GuiControl
	}
	
	try {
		XL.Run(macro)  
	} catch {
		MsgBox, 16,, Can't find the macro %A_GuiControl% in %WorkbookName%
	}
	
	return
}
