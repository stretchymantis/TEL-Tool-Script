
Open_Workbook() {                                                                                                 ; copy entire
worksheet into 'data' variable
   Global xl, wb, wb_original, ClipSaved, ExcelWorkbookPath, ActiveWorkbookPath
   ClipSaved := ClipboardAll                                      ; Save the entire clipboard
   SetBatchLines, -1

   pathResult := ResolveWorkbookPath()
   if !(IsObject(pathResult))
   {
      Clipboard := ClipSaved
      return false
   }

   file2open := pathResult.Path
   try
   {
      xl := ComObjCreate("Excel.Application")
      wb := xl.Workbooks.Open(file2open, ReadOnly:=False, Notify:=False)
   }
   catch e
   {
      Clipboard := ClipSaved
      MsgBox, 16, Excel Error, % "Failed to open '" file2open "'.`n`n" e.Message
      return false
   }

   wb_original := wb
   ActiveWorkbookPath := file2open

   if (pathResult.Save)
      SaveDatabaseSettings()

   return true
}

ResolveWorkbookPath() {
   global DatabaseSource, ExcelWorkbookPath, DefaultWorkbookPath

   if (DatabaseSource = "excel")
   {
      if (ExcelWorkbookPath != "" && FileExist(ExcelWorkbookPath))
         return { Path: ExcelWorkbookPath, Save: false }
      return PromptForWorkbookPath(ExcelWorkbookPath)
   }

   if (DefaultWorkbookPath != "" && FileExist(DefaultWorkbookPath))
      return { Path: DefaultWorkbookPath, Save: false }

   return PromptForWorkbookPath(DefaultWorkbookPath)
}

PromptForWorkbookPath(missingPath := "") {
   global DefaultWorkbookPath, ExcelWorkbookPath, DatabaseSource

   packagedPath := DefaultWorkbookPath
   message := "The 'INTEL Tool List' workbook could not be found at:`n" missingPath "`n`n"
            . "Press Yes to browse for the file, No to use the packaged data, or Cancel to stop."
   MsgBox, 3, Workbook Missing, %message%

   IfMsgBox, Cancel
      return ""

   IfMsgBox, Yes
   {
      initial := missingPath != "" ? missingPath : ExcelWorkbookPath
      FileSelectFile, userPath, 3, %initial%, Select the 'INTEL Tool List' workbook, Excel Documents (*.xlsx; *.xlsm; *.xls)
      if (userPath = "")
         return PromptForWorkbookPath("")
      if !FileExist(userPath)
      {
         MsgBox, 48, Workbook Missing, The selected file could not be found.
         return PromptForWorkbookPath(userPath)
      }
      ExcelWorkbookPath := userPath
      DatabaseSource := "excel"
      return { Path: userPath, Save: true }
   }

   if (FileExist(packagedPath))
      return { Path: packagedPath, Save: false }

   MsgBox, 16, Workbook Missing, % "The packaged workbook was not found at:`n" packagedPath
   return ""
}

InitializeSettings() {
   global SettingsIni, DatabaseSource, ExcelWorkbookPath, DefaultWorkbookPath

   if (SettingsIni = "")
      return

   if !FileExist(SettingsIni)
   {
      IniWrite, excel, %SettingsIni%, Database, Source
      IniWrite, %DefaultWorkbookPath%, %SettingsIni%, Database, ExcelPath
   }

   IniRead, dbSource, %SettingsIni%, Database, Source, excel
   IniRead, excelPath, %SettingsIni%, Database, ExcelPath, %DefaultWorkbookPath%

   if (dbSource != "internal" && dbSource != "excel")
      dbSource := "excel"

   DatabaseSource := dbSource
   ExcelWorkbookPath := excelPath
   if (ExcelWorkbookPath = "")
      ExcelWorkbookPath := DefaultWorkbookPath
}

SaveDatabaseSettings() {
   global SettingsIni, DatabaseSource, ExcelWorkbookPath
   if (SettingsIni = "")
      return
   IniWrite, %DatabaseSource%, %SettingsIni%, Database, Source
   IniWrite, %ExcelWorkbookPath%, %SettingsIni%, Database, ExcelPath
}
UpdateSettingsGuiState() {
   global DatabaseSource, ExcelWorkbookPath
   Gui, SGui:Default
   internal := DatabaseSource = "internal" ? 1 : 0
   excel := internal ? 0 : 1
   GuiControl,, DatabaseSource_Internal, %internal%
   GuiControl,, DatabaseSource_Excel, %excel%
   GuiControl,, ExcelWorkbookPath_Display, %ExcelWorkbookPath%
   if (excel)
   {
      GuiControl, Enable, ExcelWorkbookPath_Display
      GuiControl, Enable, BrowseForExcelBtn
   }
   else
   {
      GuiControl, Disable, ExcelWorkbookPath_Display
      GuiControl, Disable, BrowseForExcelBtn
   }
}

Close_Workbook() {
xl.quit
Loop, {
Process, Exist, Excel.exe
Process, Close, Excel.exe
   } Until !ErrorLevel
}
;=======================================================================

Import_Sheets() {
   Global
   For sheet in xl.ActiveWorkbook.Worksheets                                     ; Get sheetnames
   {
      FoundPos := RegExMatch(Sheet.Name, "i)^([A-Z]{3,3})$", Match)            ; Match any 3-letter sheet names
      If (FoundPos>0) {
         If (Match="BEC") ;or (Match="TGR") OR (Match="JGR")                    ; Specify here any matched sheet names to exclude
            Continue
      Matched_Sheets.Push(Match)                                  ; Matched_Sheets[] contains sheet names
      }
   ;xlSht1 := xl.ActiveWorkbook.Sheets(%2)  ;<----test this
   }
   For index, wksht in Matched_Sheets
   {
      xl.sheets(wksht).UsedRange.copy
      clipboard := RegExReplace(clipboard, "iD)\s*$")             ; Remove extra newlines at end of string (allows Excel to have empty bordered rows at the end)
      data_%wksht% := clipboard							               ; Copy all rows & columns (including headers) from the worksheet to clipboard
      StringTrimRight, data_%wksht%, data_%wksht%, 2		         ; remove final carriage returns (`r`n)
      StringUpper, data_%wksht%, data_%wksht%                     ; make uppercase
      
      Export_to_file()
   }
}

;=======================================================================
Export_to_file() {
FileAppend, % data_%wksht%, Tool_DB.csv
Return
}
;=======================================================================

Make_List() {									                        ; creates an array for each tool. Creates toolname_list and toolsn_list for the DDL
   global
   setbatchlines, -1
	For index, wksht in Matched_Sheets
   loop, parse, data_%wksht%, `r, `n							      ; 'r, 'n delimitters makes each row the delimitter
   {
		if a_index = 1												         ; a_index = 1 is header. a_index 2+ are the tools
		{
			header_%wksht%_array := strsplit(a_loopfield, a_tab)	; parse header into array
			Continue
		}
		for key_d, val_d in strsplit(a_loopfield, a_tab)		   ; Parse each row (except the header)
		{
			if (val_d = "")
            break
         if (key_d = 2)										            ; get tool name from 2nd col of row if not blank
			{
            %wksht%_list .= val_d "`n"
            tool := val_d									            ; give the name of tool the variable 'tool'
				toolname_list .= val_d "`n"	                     ; create a list of all the tool names (for DDL when 'name' is selected)
         }
         if (key_d = 3)										            ; get s/n from 3rd col of row if not blank
			{
				sn := val_d									               ; give the name of tool the variable 'sn'
				toolsn_list .= val_d "`n"						         ; create a list of all the s/ns (for DDL when s/n is selected)
				break
   		}
		}
		%tool%_array := StrSplit(a_loopfield, a_tab)			      ; a_loopfield = parsed properties of each tool. StrSplit puts in array.
   }
}

;=======================================================================
Populate_DDL() {
   GuiControl, Disable, MainDDL                          ; Prevents early-clicking of DDL from causing the list to not load
   DDL_list :=
   tool_type_quantity := tool_types.Count()  
   tally_checked = 0
   Gui, PGui:submit, nohide
   for index, key in tool_types
   {
      tally_checked += cb_%key%                          ; adds 0 to tally_checked if off. adds 1 if on (uses value of its state)
      if cb_%key% = 0                                    ; if OFF, skip the tool type
         continue
      else if cb_%key% = 1
      {
         Loop, parse, %key%_list, `n
         {
            if (R_Name1_SN2 = 1) {                                ; Populates with Tool Name if 'Name' is selected?
               DDL_list .= %a_loopfield%_array[2] "`n"
            }
            else if (R_Name1_SN2 = 2) {                           ; Populates with Serial Number if 'S/N' is selected?
               DDL_List .= %a_loopfield%_array[3] "`n"
            }
         }
      }
      if (tally_checked = tool_type_quantity)            ; if they match, then all checkboxes were checked
         guicontrol,, cb_all, 1
      else if (tally_checked < tool_type_quantity)       ; if they don't match, then not all checkboxes were checked
      {
         guicontrol,, cb_all, 0                          ; so make sure the 'all' checkbox isn't marked
      }
   }
   if (tally_checked = 0)
      DDL_List := " "
   sort, DDL_list, U
   GuiControl,, MainDDL, `n%DDL_list%					      ; ***THIS IS WHAT ACTUALLY DISPLAYS THE LIST OF ITEMS IN THE DDL***
   GuiControl, +H50, MainDDL
   GuiControl, Enable, MainDDL
   GuiControl, Focus, MainDDL
   return
}

clickbox() {                                                   ; rename to something like 'Clear_Data'                                          ; Clicking a selection in the main combobox
   setbatchlines, -1
    WinGetPos, PGuiX, PguiY, PGuiW, PGuiH, ahk_id %PGuihwnd%
   GuiControl, Show, HLV
   GuiControl, Show, PLV
  
   GuiControl, Hide, vColorBox                                 
   Gui, PGui:submit, NoHide
    if (R_Name1_SN2 = 1)                                                    ; radio button 'name' is selected. Enumerate from index 2 (which is tool name)
      array_num = 2
   else if (R_Name1_SN2 = 2) {                                             ; if radio cb 'serial' is checked, have the below loopfield array pull from index 3 (which is s/n)
      array_num = 3
   }
   ;GuiControl, Move, HLV, % "h" v_total_R_H
   ;GuiControl, Move, PLV, % "h" v_total_R_H 
  ; If LV_GetCount() > 0
  ;    Gosub Delete_Rows
   if (mainDDL_old = mainDDL)
      Return

      Loop, parse, toolname_list, `n
      {
         if (a_loopfield = mainDDL)     ; evaluate if combobox matches a tool name
         {
            Selected_Tool := a_loopfield
            gosub Delete_Rows
            gosub Add_Rows
            Return
         }
      Continue
      }
      If (mainDDL = "")        ;If DDL is blank, delete the rows and that's it
      {
         Gosub Delete_Rows
         PGuiH := 96
         Winset, Region, % "w" PGuiW " h" PGuiH " 0-0 R15-15", ahk_id %PGuihwnd%    ; make bottom of window rounded after pgui unrolls
         Return
      }
      Else if !instr(DDL_List, mainDDL)            ; evaluate if combobox matches a tool name
      {
         ; Msgbox, % a_loopfield_array[array_num] " does not exist in the spreadsheet."
         Return
      }
      Else                  ;If DDL *not* blank, Delete rows then add new rows for selected tool
      {
      Return
      }
   
   Delete_Rows:   ; ***** DELETE ROWS *****
   Loop % LV_GetCount()          ; Delete pre-existing rows, if any (ArrayCount will be empty if 1st time running)
   {
     ; msgbox, %a_index%
      ;Sleep 50
      Gui, PGui:ListView, hlv    ; make hlv (header) column default
      LV_Delete(1)               ; delete text from hlv (header) column
      Gui, PGui:ListView, plv    ; make plv (properties) column default
      LV_Delete(1)               ; delete text from plv (properties) column
      ;v_total_R_H -= RowHeight   ; v_total_r_h is the cumulative pixel height of rows added * height of each row (19)
      ;GuiControl, Move, HLV, % "h" v_total_R_H  ; reduce HLV column (removes underline)
      ;GuiControl, Move, PLV, % "h" v_total_R_H  ; reduce PLV column (removes underline)
      ;WinGetPos, PGuiX, PguiY, PGuiW, PGuiH, ahk_id %PGuihwnd%
      ;Winset, Region, % "w" PGuiW " h" PGuiH-RowHeight " 0-0 R15-15", ahk_id %PGuihwnd%    ; make bottom of window rounded after pgui unrolls
      ;Gui, PGui:Show, % "w" PGuiW " h" PGuiH-RowHeight
      ;WinGetPos, PGuiX, PguiY, PGuiW, PGuiH, ahk_id %PGuihwnd%
      }
  ; Gui, PGui:Show, H150
   Return
   
   Default_View:
   Loop, parse, toolname_list, `n
   {
      if (%a_loopfield% = mainDDL) {             ; evaluate if combobox matches a tool name
   ;  if (%a_loopfield%_array[array_num] = mainDDL) {             ; evaluate if combobox matches a tool name
         v_total_R_H := 19
         val := substr(a_loopfield, 1, 3)                         ; get 1st three letters of toolname
         ; ++++ COUNT NUMBER OF HLV AND PLV ROWS ++++
         for key_h, val_h in header_%val%_array                   ; loop thru header to add to listview
         {
            if (header_%val%_array[key_h]="") OR (%a_loopfield%_array[a_index]="") ; if both header and property values are blank (ie spreadsheet contained blank row), then skip it
            {
               continue
            }
            sleep 25
      Gui, PGui:ListView, hlv							         ; make the HEADER the default (hlv)}
      LV_Add(, header_%val%_array[key_h])			         ; populates header's ListView (hlv) (%val% will be either TGR or JGR depending on what was selected
      Gui, PGui:ListView, plv							         ; make PROPERTIES default (plv)}
      LV_Add(, %a_loopfield%_array[a_index])			      ; populates the tool's ListView (plv
      Gui, PGui:Show, Autosize
         }
      }
   }
}

   Add_Rows:
   ; Loop, parse, toolname_list, `n
   ; {
      v_total_R_H := 19                                        ; cumulative pixel height of rows added; this starts the 'Add_Rows' process with 1 row visible
      val := substr(Selected_Tool, 1, 3)                         ; get 1st three letters of toolname
      ; ++++ COUNT NUMBER OF HLV AND PLV ROWS ++++
      for key_h, val_h in header_%val%_array                   ; loop thru header to add to listview
      {
         if (header_%val%_array[key_h]="") OR (%a_loopfield%_array[a_index]="")      ; if header or property values are blank (ie spreadsheet contained blank row), then skip it
         {
            continue
         }
         ;sleep 25
         ; ++++ ADD TO ROWS ++++
         v_total_R_H += RowHeight                          ; v_total_r_h is the cumulative pixel height of rows added * height of each row (19)
         GuiControl, Move, HLV, % "h" v_total_R_H
         GuiControl, Move, PLV, % "h" v_total_R_H
         Gui, PGui:ListView, hlv							         ; make the HEADER the default (hlv)}
         LV_Add(, header_%val%_array[key_h])			         ; populates header's ListView (hlv) (%val% will be either TGR or JGR depending on what was selected
         Gui, PGui:ListView, plv							         ; make PROPERTIES default (plv)}
         LV_Add(, %Selected_Tool%_array[a_index])			      ; populates the tool's ListView (plv) with the selected tool's properties into the ListView
         ;Gosub Cu-NonCu_Colored_Text                        ; Uncomment to use: Changes color of the Cu/Non-Cu text
         WinGetPos, PGuiX, PguiY, PGuiW, PGuiH, ahk_id %PGuihwnd%
         Winset, Region, % "w" PGuiW " h" PGuiH " 0-0 R15-15", ahk_id %PGuihwnd%    ; make bottom of window rounded after pgui unrolls
         ;Gui, PGui:Show, Autosize
         ;Gui, PGui:Show, % "w" PGuiW " h" PGuiH+RowHeight
         continue
         Gui, PGui:Show, Autosize
         ; return
      }
   mainDDL_old := mainDDL
   return
   ; }

;=======================================================================

WM_MOUSEMOVE() {
    static CurrControl, PrevControl, _TT  ; _TT is kept blank for use by the ToolTip command below.
    CurrControl := A_GuiControl
    If (CurrControl <> PrevControl and not InStr(CurrControl, " "))
    {
        ToolTip  ; Turn off any previous tooltip.
        SetTimer, Show_Hover_Help, 1000
        PrevControl := CurrControl
    }
    return

    Show_Hover_Help:
    SetTimer, Show_Hover_Help, Off
    ToolTip % %CurrControl%_TT
    SetTimer, Hide_Hover_Help, 3000
    return

    Hide_Hover_Help:
    SetTimer, Hide_Hover_Help, Off
    ToolTip
    return
}

LVGetItemHeight(hwnd_lv1) {                                    ; allows resizing the ListView height w/out scrollbars
   Static LVM_GETITEMRECT := (4096 + 14)
   VarSetCapacity(RECT, 16, 0)
   SendMessage, LVM_GETITEMRECT, 0, &RECT, , % "ahk_id " . hwnd_lv1
   Return NumGet(RECT, 12, "Int") - NumGet(RECT, 4, "Int")
}

GuiGetSize( ByRef W, ByRef H, GuiID=1 ) {
	Gui %GuiID%:+LastFoundExist
	IfWinExist
	{
		VarSetCapacity( rect, 16, 0 )
		DllCall("GetClientRect", uint, MyGuiHWND := WinExist(), uint, &rect )
		W := NumGet( rect, 8, "int" )
		H := NumGet( rect, 12, "int" )
	}
}

GuiGetPos( ByRef X, ByRef Y, ByRef W, ByRef H, GuiID=1 ) {
	Gui %GuiID%:+LastFoundExist
	IfWinExist
	{
		WinGetPos X, Y
		VarSetCapacity( rect, 16, 0 )
		DllCall("GetClientRect", uint, MyGuiHWND := WinExist(), uint, &rect )
		W := NumGet( rect, 8, "int" )
		H := NumGet( rect, 12, "int" )
	}
}

;=======================================================================
IsVar(__var_, __DisplayText_:=0) {
	DetectHiddenWindows, On
	WinMenuSelectItem, ahk_id %A_ScriptHwnd%, , View, 2&
	WinHide, ahk_id %A_ScriptHwnd%
	WinGetText, __varlist_, ahk_id %A_ScriptHwnd%                  ;ahk_class Edit1
	WinHide, ahk_id %A_ScriptHwnd%
	if InStr(__varlist_, "`n" __var_ "[") OR InStr(__varlist_, "`n" __var_ ": Object")Return 1if __DisplayText_
		MsgBox %__varlist_%
	Return 0
}



/*  TEMPORARILY DISABLED FOR MAP.AHK INSERTED CODE
*/
class Color {
  __New(aRGB = 0x000000) {
    this.RGB := aRGB
  }
  __Get(aName) {
    if (aName = "R")
      return (this.RGB >> 16) & 255
    if (aName = "G")
      return (this.RGB >> 8) & 255
    if (aName = "B")
      return this.RGB & 255
    if (aName = "hex")
    {
      format_setting := A_FormatInteger
      SetFormat, IntegerFast, h
      hex := SubStr(this.RGB + 0, 3)
      SetFormat, IntegerFast, %format_setting%
      while StrLen(hex) < 6
	hex := "0" . hex
      return, "0x" . hex
    }
  }
  __Set(aName, aValue) {
    if aName in R,G,B
    {
      aValue &= 255
      if      (aName = "R")
	this.RGB := (aValue << 16) | (this.RGB & ~0xff0000)
      else if (aName = "G")
	this.RGB := (aValue << 8)  | (this.RGB & ~0x00ff00)
      else  ; (aName = "B")
	this.RGB :=  aValue        | (this.RGB & ~0x0000ff)
      return aValue
    }
  }
}

;=======================================================================
MouseIsOver(WinTitle) {
 MouseGetPos,,, Win
 return WinExist(WinTitle . " ahk_id " . Win)
                                                                                                                                                                                 }


;================================LABELS=================================
;Return

Toggle_All:                                                       ; TOGGLES ALL CHECKBOXES
   Gui, PGui:submit, nohide
   if cb_ALL = 1                                                  ; if CB_ALL checkbox is ON
   {
      for index, key in tool_types                                ; tool_types=CAT,TGR,JGR,etc...
      {
         guicontrol,, cb_%key%, 1                                 ; turn checkbox ON
      }
   }
   else if cb_All = 0
   {
      for index, key in tool_types
      guicontrol,, cb_%key%, 0
   }
   populate_ddl()
   return

R_Name_Serial:                                              ;RADIO BUTTONS 'NAME' AND 'S/N' BEHAVIOR
   Gui, PGui:submit, nohide
   if (R_Name1_SN2 = 1)
   {     
      DDL_list := toolname_list
      Guicontrol,, gb1, Tool Name
   }
   else if (R_Name1_SN2 = 2) {
      DDL_list := toolsn_list
      Guicontrol,, gb1, Serial Number
   }
   else
      Msgbox "Could not determine radio button state"
   populate_ddl()
   clickbox()
   return

Check:
   LV_Modify(1, "-Focus -Select")
   If !LV_GetNext(0)
   LV_Modify(2, "Select")
   If !LV_GetNext(0, "F")
   return

Unselect:
   msgbox, clicked
   LV_Modify(0, -focus)
   return

E1:
	Gui,Submit,NoHide
	sitesArr := StrSplit(sites, "|")
	newArr   := []
	newStr   := ""
	for k,v in sitesArr
	{
		if InStr(v, Search, false)>0
		{
			newArr.push(v)
		}
	}
	for k,v in newArr
		newStr .= "|" v
	GuiControl, , Site, % newStr
	GuiControl, Choose, site, 1
   Return

Cu-NonCu_Colored_Text:
   if (%a_loopfield%_array[a_index] = "CU" || %a_loopfield%_array[a_index] = "NON-CU") {          ; check if Cu or Non-Cu, then create text overlay w/ font color
      GuiControlGet, PLV, Pos                         ; get position of PLV ListView
      GuiControl, Move, vColorBox, % "x" PLVX+6 " y" PLVY+v_Total_R_H-RowHeight+2               ; move the Cu/Non-Cu text field over the listview one
      GuiControl, Show, vColorBox
      if (%a_loopfield%_array[a_index] = "CU") {
         Gui, font, S9 w1000 c996515, Segoe UI
         GuiControl, Font, vColorBox
         GuiControl,, vColorBox, CU
      }
      else if (%a_loopfield%_array[a_index] = "NON-CU") {
         Gui, font, S9 w1000 cGreen, Segoe UI
         GuiControl, Font, vColorBox
         GuiControl,, vColorBox, NON-CU
      }
      LV_Modify(a_index,, "")
   }
   Gui, font, S9 w1000 cWhite, Segoe UI
   Gui, Show, Autosize
   Return

+Control::
ESCaped := 0											               ; Press {CTRL} twice for ToolTip; three times for TEL Info Tool
if (control_presses > 0) 							               ; SetTimer already started, so we log the keypress instead.
{
   control_presses += 1
   return
}
control_presses := 1                                        ; Otherwise, this is the first press of a new series. Set count to 1 and start the timer
SetTimer, KeyControl, -800 						               ; Wait for more presses within specified time.
return

; OCR CHECK UNDER MOUSE
KeyControl:                                                 ; ======= ALL GUITIP STUFF SHOULD GO IN THIS LABEL ========
if (control_presses = 2) { 					                  ; The key was pressed twice.
   ;if WinExist("ahk_id" hGuiTip1) {                         ; stop a currently running GuiTip window (still need to get rid of that weird circle (vFadingTextControl1?)
      ; Gui, %hGuiTip1%: Destroy
      control_presses := 0
      Gosub Remove_Tooltip
 ;  }
   OCR_Text := OCR_Area()                                   ; OCR the cursor location
   OCR_Text := OCR_CleanFix()                               ; Fix accenting & incorrect characters (may result in null if no tool found)
   StringUpper, OCR_Text, OCR_Text
   if OCR_Match() = 1 {                                     ; Determine if there's a match to an existing tool
      if InStr(toolname_list, OCR_Text) && (OCR_Text != "") { ; check that captured string exists in tool_list and isn't blank
         ; TOOLTIP DISPLAY FROM OCR RESULTS
         SN := %OCR_Text%_array[3]  ; Use 'SN' variable for serial number in below txt bc BeatifulToolTip can't seem to use double variable references
         Text= Press ESC to close`n`n%OCR_Text%:  %SN%
         Text2= Press ESC to close
         Style=Style2
         if Is_ToolTip_Shown = 1    If tooltip is currently being shown, remove it
            Gosub Remove_Tooltip

         Gosub Show_ToolTip      ; TGR401
         Return
      }         
      Show_ToolTip:
      ; fadein
      for k, v in [5,10,15,35,55,65,75,85,95,105,115,125,135,145,155,165,175,185,195,200,205,215,220,225,230,235,240,245,250,255]
      {
         btt(Text,,,,Style,{DistanceBetweenMouseYAndToolTip:-60, Transparent:v})
         Sleep, 30
      }
      SetTimer, ToolTip_Follow_Mouse, 10
      Is_ToolTip_Shown := 1
      return
      
      ToolTip_Follow_Mouse:
      btt(Text,,,,Style,{DistanceBetweenMouseYAndToolTip:-60})
      return

   }                                                        ; OnMessage(0x201, Func("FOLLOWMOUSE").Bind(hGui, GuiHomeX, GuiHomeY, Period, Step)) ; Enable to require a click to follow mouse
   else                   ; tgr401 hello                         ; no match
   {   
      Text= Nothing found.`nTip: To get info about a tool,`nplace cursor over a tool name`nand press the hotkey.
      Style= Style9
      Gosub Show_ToolTip2
      control_presses := 0 					                  ; regardless which was triggered, reset count to prepare for the next series of presses
   sleep, 4000
   Gosub Remove_ToolTip
   }
}
else if (control_presses > 2)	{						            ; CTRL was pressed 3 or more times;
   Gui, Pgui:Default
   Gui, +LastFound
   Gui, show
}
control_presses := 0 									            ; regardless which was triggered, reset count to prepare for the next series of presses
return

      Show_ToolTip2:
      ; fadein
      for k, v in [5,10,15,35,55,65,75,85,95,105,115,125,135,145,155,165,175,185,195,200,205,215,220,225,230,235,240,245,250,255]
      {
         btt(Text,,,,Style,{DistanceBetweenMouseYAndToolTip:-60, Transparent:v})
         Sleep, 30
      }
      SetTimer, ToolTip_Follow_Mouse, 10
      Is_ToolTip_Shown := 1
      return


Remove_ToolTip:
; fadeout TGR401
If (Is_ToolTip_Shown = 0)
   Return
SetTimer, ToolTip_Follow_Mouse, Off
for k, v in [240,220,200,180,160,140,120,100,80,60,40,20,0]
{
   btt(Text,,,,Style,{DistanceBetweenMouseYAndToolTip:-60, Transparent:v})
   Sleep, 30
}
Is_ToolTip_Shown := 0
BTT()
Return

PLV_Copy:											                    ; copy clicked control to clipboard
   Gui, PGui:ListView, plv
   if (A_GuiEvent = "Normal")
   {
      LV_GetText(sel_prop, A_EventInfo)                       ; sel_prop is the selected property
      Clipboard := sel_prop
      LV_Modify(A_EventInfo, "-Select")
      LV_Modify(A_EventInfo, "-Focus")
      ToolTip, % "'" sel_prop "' copied to clipboard"
      SetTimer, RemoveToolTip, -2000
      ;SetTimer, Hide_Hover_Help, -2000 TEMPORARY FOR MAP.AHK
   }
   else if (A_GuiEvent = "D")
   {
      LV_Modify(A_EventInfo, "-Select")
      LV_Modify(A_EventInfo, "-Focus")
   }
   else if (A_GuiEvent = "RightClick")
   {
      LV_Modify(A_EventInfo, "-Select")
      LV_Modify(A_EventInfo, "-Focus")
   }
return

HLV_click:
   Gui, PGui:ListView, hlv
   if (A_GuiEvent = "Normal")
   {
      LV_Modify(A_EventInfo, "-Select")
      LV_Modify(A_EventInfo, "-Focus")
   }
     else if (A_GuiEvent = "D")
   {
      LV_Modify(A_EventInfo, "-Select")
      LV_Modify(A_EventInfo, "-Focus")
   }
else if (A_GuiEvent = "RightClick")
   {
      LV_Modify(A_EventInfo, "-Select")
      LV_Modify(A_EventInfo, "-Focus")
   }
return

UIMove:
PostMessage, 0xA1, 2,,, A
Return

PGuiGuiContextMenu:
Menu, MainMenu, Show
return

MenuSelection_Help:
Gui, Help:New,, Help
Gui, Add, Text,,
Gui, Show
return

MenuSelection_Settings:
   UpdateSettingsGuiState()
Gui, SGui:Show
return

DatabaseSourceChange:
   Gui, SGui:Submit, NoHide
   if (DatabaseSource_Internal)
      DatabaseSource := "internal"
   else if (DatabaseSource_Excel)
      DatabaseSource := "excel"
   SaveDatabaseSettings()
   UpdateSettingsGuiState()
return

BrowseForExcel:
   Gui, SGui:+OwnDialogs
   initial := ExcelWorkbookPath
   FileSelectFile, newPath, 3, %initial%, Select the 'INTEL Tool List' workbook, Excel Documents (*.xlsx; *.xlsm; *.xls)
   if (newPath = "")
      return
   if !FileExist(newPath)
   {
      MsgBox, 48, Workbook Missing, The selected file could not be found.
      return
   }
   ExcelWorkbookPath := newPath
   DatabaseSource := "excel"
   SaveDatabaseSettings()
   UpdateSettingsGuiState()
return

SGuiClose:
SGuiEscape:
   Gui, SGui:Hide
return

MenuSelection_About:
msgbox, Under Development
return


MenuSelection_Quit:
ExitApp

MenuHandler:
Gui, PGui:+Owndialogs
Gui, PGui:+Disabled
MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
Gui, PGui:-Disabled
return

AOTsub:											                     ; always on top option
Gui, PGui:submit, nohide
WinSet, AlwaysOnTop, Toggle
return

PGuiShow:
Gui, PGui:Show
return

RemoveToolTip:
ToolTip
return


;********************************************************************************************************************************************************************************************************
;********************************************************************************************************************************************************************************************************
;********************************************************************************************************************************************************************************************************
;********************************************************************************************************************************************************************************************************
/*
class LWButtons	{
	;class: Layered Window Buttons
	__New( obj := "" ){
		This._SetDefaults()
		This._UpdateDefaults( obj )
		This._CreateControl()
		This._DrawButton()
	}
	_SetDefaults(){
		
		This.Hwnd := ""
		This.Button := ""
		
		This.X := 10
		This.Y := 10 
		This.W := 10
		This.H := 10
		
		This.RestColor := "0xFF3399FF"
		This.PressedColor := "0xFF336699"
		This.MovingColor := "0xFFff0000"
		
		This.ButtonText := "Button"
		This.Font := "Comic Sans MS"
		This.FontSize := 16
		This.FontColor := "0xFF000000"
		This.FontOptions := " Center vCenter Bold "
		
		This.Roundness := 10
		This.Label := ""
		This.WindowOptions := " -DPIScale +AlwaysOnTop +ToolWindow "
		
		This.Pressed := 0
		
		This.ClickBind := This._ButtonClick.Bind( This )
		;~ OnMessage( 0x201 , This._ButtonClick.Bind( This ) ) ; I could have gone this route, but I like the path I took more.
	}
	_UpdateDefaults( obj := "" ){
		local k , v 
		for k, v in obj
			This[ k ] := obj[ k ]
	}
	_CreateControl(){
		local hwnd , bd := This.ClickBind
		This.Button := New PopUpWindow( { AutoShow: 1 , X: This.X , Y: This.Y , W: This.W , H: This.H , Options: This.WindowOptions } )
		This.Hwnd := This.Button.Hwnd
		Gui, % This.Button.Hwnd ":Add" , Text , % "x" 0 " y" 0 " w" This.W " h" This.H " hwndhwnd"
		GuiControl, % This.Button.Hwnd ":+G" , % hwnd , % bd
	}
	_ButtonClick(){
		if( GetKeyState( "Shift" ) ){
			This.Pressed := 2
			This._DrawButton()
			PostMessage, 0xA1, 2
			While( GetKeyState( "LButton" ) )
				sleep, 60
			WinGetPos, x, y ,,, % "ahk_id " This.Button.Hwnd
			This.Button.UpdateSettings( { X: x , Y: y } )
			This.Pressed := 0
			This._DrawButton()
		}else{
			This.Pressed := 1
			This._DrawButton()
			While( GetKeyState( "LButton" ) )
				sleep, 60
			This.Pressed := 0
			This._DrawButton()
			MouseGetPos,,, win
			if( win = This.Button.Hwnd ){
				if( isObject( This.Label ) )
					Try
						This.Label.Call()
				else 
					Try
						GoSub, % This.Label
			}
		}
		
	}
	_DrawButton(){
		This.Button.ClearWindow()
		This._ButtonGraphics()
		This.Button.UpdateWindow()
	}
	_ButtonGraphics(){
		Brush := Gdip_BrushCreateSolid( ( This.Pressed = 0 ) ? ( This.RestColor ) : ( This.Pressed = 1 ) ? ( This.PressedColor ) : ( This.MovingColor )  ) , Gdip_FillRoundedRectangle( This.Button.G, Brush , 1 , 1 , This.W - 2 , This.H - 2 , This.Roundness ) , Gdip_DeleteBrush( Brush )
		Brush := Gdip_BrushCreateSolid( This.FontColor ) , Gdip_TextToGraphics( This.Button.G , This.ButtonText , "s" This.FontSize " c" Brush " " This.FontOptions " x0 y0" , This.Font, This.W , This.H ) , Gdip_DeleteBrush( Brush )
	}
	Move( x := 10 , y := 10 ){
		This.Button.UpdateSettings( { X: x , Y: y } )
		Gui, % This.Button.Hwnd ":Show", % "x" x " y" y " NA"
	}
	Hide(){
		Gui, % This.Button.Hwnd ":Hide" 
		;~ This.Button.ClearWindow( 1 ) ; <<---------------------  This is how I would normally "Hide" a layered window.
	}
	Show(){
		Gui, % This.Button.Hwnd ":Show", NA 
		;~ This._DrawButton() ; <<------ The way I would normally do it.
	}
	Delete(){
		This.Button.DeleteWindow()
	}
}

*/

; separator
;Menu, Tray, NoStandard
;menu, tray, add, Restore Window, RestoreWindow
;menu, tray, add 
;menu, tray, add, Close Program, CloseProgram
;menu, tray, Default, Restore Window

return

;-------------------------------------------------------------------------------
WM_LBUTTONDOWN() { ; move window
;-------------------------------------------------------------------------------
    PostMessage, 0xA1, 2 ; WM_NCLBUTTONDOWN
}

RestoreWindow:
	WinShow, My_Windows_Title
	return

CloseProgram:
	Exitapp
	return

#IF Is_ToolTip_Shown = 1
ESC:: Goto Remove_ToolTip
#If    ; TGR401



 
