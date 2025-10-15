SetBatchLines, -1
ClipSaved := ClipboardAll 
StartedFromMainScript := True
;#Include FFToolTip.ahk
;#Include Window_Dragging.ahk

#Include OCR.ahk
;#Include ThrowWindows.ahk

; RegEx Fla */vor: PCRE 8.30
; In RegExBuddy:  pcre830 4.0.0 PCRE 8.30–8.33 UTF-8
;                 pcre830 4.0.0 PCRE 8.30–8.33

; TO DO BEFORE COMPILING:
;
;  -remove any unnecessary global declarations
;  -uncomment the tool tray icon
;  -uncheck Always on Top
;  -check script with #Warn ON
;  -immediately deselect  a listview itemafter clicking
;  -create 'copied' balloon when clicking a listview item
;  -fix Excel file being stuck open sometimes
;  -Add CAT sheet back in and set it up to be read correctly
;  -color cells of cu/noncu (animate like a stripe?)
;  -add ability to show by FAB
;  -add ability to view all tools as a list from a sidebar that populates main window
;  -add ability to 'dock' the main window to the screen edge and allow it to be moved and to pop out

; TO DO & Ideas (GENERAL)
;  -Use select -> copy 1st before OCR attempt

; [ADD TO APP]

; MSDS
; The More You Know
; Spec sheet
; Piping diagram symbols

; [TEL LOGO COLORS]
; Excel map cell size: Width=70px ; Height=25px
; Grey:                    #808080
; Blue:   RGB=0,172,235;  #00ACEB
; Green:  RGB=88,181,48; #58b530
; Copper: #523019 #734621

;http://autohotkey.com/board/topic/29449-gdi-standard-library-145-by-tic/

;HOTKEYS:
;/*************************************
;F1: load new image
;esc: exit
;lbutton: drag image
;space: zoom large image to fit screen
;rbutton: reload image
;up/wheelup: scale up
;down/wheeldown: scale down
;left/xbutton1: rotate left
;right/xbutton2: rotate right

; ***** SETTINGS *****  
;       ***** ToolTip Mouse Hover Descriptions *****
mainddl_TT  := "Select or Type a Tool Name or S/N"
AOT_TT      := "Always keep this window on top of other windows_NewEnum()`nTIP: When selected, you can change the opacity of the window in Settings."
MAP_TT      := "Show the map"

step  := 0.1  ;set zoom step in px (absolute) OR  percent (relative, eg 0.1 = 10%)
angle := 90   ;set angle step in degrees
lres  := true ;resample @ half res (faster for large images)

DetectHiddenWindows, On
#Include Gdip_All.ahk
#Include BTT.ahk
;#Include ThrowWindows.ahk
;#Include OCR_Del.ahk
;#Include LV_EX.ahk

#NoEnv						      			                  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance On
#Persistent
;#Warn				      					                  ; Enable warnings to assist with detecting common errors.
SendMode Input    							                  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%				                  ; Ensures a consistent starting directory

;gdip init
pToken := Gdip_Startup()
If !pToken
{
	MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
	ExitApp
}

; [GLOBAL VARIABLE DECLARATIONS]
Global DDL_List
global main :=, x :=, y :=
Global maptextwidth := 100				;Move back when done testing
Global maptextheight := 300				;Move back when done testing
Global MapBackW := 200
Global zMapBackW := MapBackW
Global MapBackH := 200
Global zMapBackH := MapBackH
Global MoveGui, Timer, hGuiTip1, hGuiTip2, hGuiNum, ESCaped, master_list, array_key, array_value, x_gb1, y_gb1, w_gb1, h_gb1, x_gb2, y_gb2, w_gb2, h_gb2
global toolserial_array := Array()
global ToolTip_Toggle := false                           ; used as a tooltip toggle to determine if tooltip is active or not.
Global data := ""						                        ; Variable for holding Clipboard contents
Global mainDDL := ""					                        ; What is actually selected in the DDL box
Global mainDDL_header, MAP, AOT, wksht
Global sheet_tab := ()
Global tool_list := ""
Global mouseX, mouseY
Global OffsetX, OffsetY, GuiTipX, GuiTipY, GuiTipW							                        ; GuiTip width, GuiTipH								                     ; GuiTip height
Global OCR_Text                            ; TGR401
Global FadingTextControl1
Global FadingText1 := "Cancel = ESC"                     ; 1st fading text line at bottom of tooltip
Global FadingText2 := "CTRLx2 = more"                    ; 2nd line
Global guiname := ""
Global matchedpatternABC
Global matchedpattern123
Global GuiW, GuiH
Global hGui
Global sGuiName
Global RowHeight
Global PLV, HLV, PLVH
Global MAP, MAPX, MAPY, MAPW, MAPH, AOT, AOTX, AOTY, AOTW, AOTH
Global vColorBox, vColorBoxX, vColorBoxY, vColorBoxW, vColorBoxH, vColorBoxCount := 0
Global hwnd_lv1, hwnd_pic1, hwnd_pGui, hwnd_aot, hwnd_lv1, hwnd_lv2
Global x_aot, y_aot, w_aot, h_aot, x_pic1, y_pic1, w_pic1, h_pic1, y_aot_reset, y_pic1_reset
Global v_total_R_H
Global R1W, R2W, R1, R2
Global R_Name1_SN2
Global HLV
Global Rows := 0
Global skip_first
Global ToolName_List, ToolSN_List
Global sheet_tab := ["TGR", "JGR"]
Global tool_types := ["CAT", "JGR", "TGR", "LEO"]                              ; Currently Manually added; sheet names in Tool List Excel file to pull data from; add/delete affects tool_quantity and tool_types
Global Matched_Sheets :=[]
Global vtemp, array_num, ArrayCount, ArrayCount2
Global MyInstance, MapW_Section, MapH_Section, MapX_Orig
Global ScaleResize ; := 21
Global CenterX :=, CenterY :=
Global maptextwidth :=, maptextheight :=
Global PGuiH, PGuiW, PGuiX, PguiY
Global c :=
Global Selected_Tool
Global MainDDL_old

Global SettingsIni := A_ScriptDir "\settings.ini"
Global DefaultWorkbookPath := A_ScriptDir "\INTEL Tool List Rev2.xlsx"
Global DatabaseSource := ""
Global ExcelWorkbookPath := ""
Global ActiveWorkbookPath := ""

InitializeSettings()

background_color := 0x00aceb                             ; fade color of text at bottom of tooltip
text_color := 0x000000                                   ; initial color of text at bottom of tooltip
fade := new Text_fader(text_color, background_color)
index_toggle := "Closed"				                     ; whether settings menu is open or closed
hwnd_ddl1 := ""					      	                  ; HWND of the DropDownList
tool := ""
settings_color := "00aceb"
clipcontents = 0	                                       ; Initialize variable
toolvalue = 0                    		                  ; Initialize variable
toggler := 0
label_tool_sn := "Tool"
Step := 15                            	                  ; follow mouse steps (higher =faster)
Period := 10 
Global PGuihwnd :=                          	                  ; follow mouse rest period (higher = smoother)
Global winhwnd :=

; [SETTINGS GUI]
Gui, SGui:New,, Settings Menu
Gui, Add, Checkbox, xm+5 y24 w100 h30 , Run at Startup
Gui, Add, Checkbox, xm+5 y224 w270 h30, Activate when cursor hovers over Tool Name or S/N
Gui, Add, Checkbox, xm+5 y154 w200 h30, Activate with shortcut key CTRL+]
Gui, Add, Checkbox, x252 y419 w0 h0   , CheckBox
Gui, Add, Text,     x117 y374 w40 h0  , Keyboard Shortcut:
Gui, Add, Text,     x145 y362 w110 h30, Internal database date:
Gui, Add, Radio,    xm+5 y354 w130 h30 vDatabaseSource_Internal gDatabaseSourceChange, Use internal database
Gui, Add, Radio,    xm+5 y384 w240 h30 vDatabaseSource_Excel gDatabaseSourceChange, Build database from 'INTEL Tool List' Excel file
Gui, Add, Text,     xm+21 y414 w80 h23, Excel file:
Gui, Add, Edit,     x+5 yp-3 w165 vExcelWorkbookPath_Display +ReadOnly
Gui, Add, Button,   x+5 yp+1 w70 h23 gBrowseForExcel vBrowseForExcelBtn, Browse...
Gui, Add, Checkbox, xm+5 y124 w100 h30, Always on Top
Gui, Add, Text,     xm+5 y94 w70 h30  , Transparency
Gui, Add, Checkbox, xm+5 y254 w180 h30, Deactivate balloon with ESC key
Gui, Add, GroupBox, xm y3 w282 h60    , General Settings
Gui, Add, GroupBox, xm y73 w285 h119  , TEL Info Tool
Gui, Add, GroupBox, xm y204 w290 h120 , Balloon
Gui, Add, GroupBox, xm y334 w290 h120 , Database
Gui, Add, Text, xm21 y284 w250 h30    , Note: Balloon disappears after 60 seconds
Gui, Add, Slider, x85 y90 w100 h30

UpdateSettingsGuiState()

; =============== TEL INFO TOOL GUI ==================
Margin       := 10        ; space between two group control boxes
h_cb         := 23        ; the height of the checkbox group control (increases both group controls)

; ***** MAIN GUI *****
groupbox_width := 150
Gui, PGui:New, +alwaysontop +Lastfound +hwndPGuihwnd -dpiscale  +delimiter`n +0x00040000 -caption -resize, TEL Info Tool
Gui, Color, %settings_color%
Gui, Margin, 15, 15
Gui, Font, S8 W400
Gui, Add, GroupBox, % "w" groupbox_width " h" 2*(h_cb+12)" +hwndhwnd_gb1 vgb1 Section"                 ; 26 automatically getting added to y for title bar I think
GuiControlGet, GB1, Pos
Gui, Add, ComboBox, xp+10 yp+18 w86 vmainddl +hwndhwnd_ddl1 gclickbox Sort              ; gclickbox runs when new item from DDL is selected.
Gui, Add, Radio,    xs+10 y+5 checked gr_name_serial vR_Name1_SN2 +hwndhwnd_r1, Name    
; vr_name_serial: 1=Tool Name; 2=S/N
Gui, Add, Radio,    x+0 yp+0 gr_name_serial +hwndhwnd_r2, S/N
Gui, PGui:submit, nohide
If (R_Name1_SN2 = 1)
    R_Name1_SN2_Label = Tool Name
Else
    R_Name1_SN2_Label = Serial Number
GuiControl,, gb1, % R_Name1_SN2_Label
Gui, Add, Picture,  x+0 y+-40 h35 w-1 vMAP +hwndhwnd_pic1 gShow_Map, map-icon_nomap.png     ;default size of map_icon = w41xh36
Gui, Add, GroupBox, % "xs+" GB1W+20 " ys" " w" groupbox_width+30 " h" 2*(h_cb+12) " vgb2 +hwndhwnd_gb2 Section", Show/Hide
Gui, Add, CheckBox, xp+15 yp+20 gpopulate_ddl vcb_CAT -checked section,  CAT
Gui, Add, CheckBox, x+0 yp+0 gpopulate_ddl vcb_TGR checked,              TGR
Gui, Add, CheckBox, x+0 yp +hwndhwnd_cb1 vcb_JGR gpopulate_ddl checked,  JGR
Gui, Add, CheckBox, xs y+10 vcb_LEO gpopulate_ddl,                       LEO
;Gui, Add, CheckBox, x+0 yp vcb_SRT gpopulate_ddl,                        SRT
Gui, Add, CheckBox, x+0 yp vcb_ALL gtoggle_all section,                  ALL
Gui, Font, S9 w1000 cWhite, Segoe UI
Gui, Add, ListView, % "xm w105 R1 vhlv c3A3B3C +hwndhwnd_lv1 gHLV_click Section AltSubmit +0x4000000 LV_0X20 -E0x200 grid -border -multi -hdr Background00aceb", header1							;E0x200 removes the thin border around the whole control
Gui, Add, ListView, % "x+m yp R1 w200 vplv +hwndhwnd_lv2 gPLV_copy AltSubmit +0x4000000 CWhite LV0x8000 LV_0X20 -E0x200 grid -border -multi -hdr Background00aceb", prop	;E0x200 removes the thin border around the whole control
GuiControl, Hide, HLV
GuiControl, Hide, PLV
Gui, PGui:ListView, hlv
LV_add()                                        ; Just to determine RowHeight in pixels. Added so LVGetItemHeight has a row to measure
RowHeight := LVGetItemHeight(hwnd_lv1)          ; measure the height of the row
LV_delete()
;skip_first := 1                                ; used to skip first y-increase iteration for aot & map.
Gui, font,        S12
gui, add,   text, % "x" groupbox_width + 140 " y" -6 gPMinimize, __
gui, add,   text, % "x" groupbox_width + 170 " y" -2 gPExit, x
Gui, font,        S8 W400
Gui, Add,   Text, VvcolorBox, TESTTESTTEST
GuiControl, Hide, vColorBox

; END MAIN GUI PART
;  [CONTEXT MENU]
Menu, MainMenu, Add, Help, MenuSelection_Help					               ; create the popup menu by adding some items to it.
Menu, MainMenu, Add, Settings, MenuSelection_Settings
Menu, MainMenu, Add, About, MenuSelection_About  				               ; add a separator line
Menu, MainMenu, Add, Quit, MenuSelection_Quit

; [TRAY MENU]
;Menu, Tray, Icon, TEL_Icon.ico
;Menu, Tray, NoStandard
Menu, Tray, Add, Open TEL Info Tool, PGuiShow
Menu, Tray, Add
Menu, Tray, Add, Help, MenuSelection_Help
Menu, Tray, Add, Settings, MenuSelection_Settings
Menu, Tray, Add, About, MenuSelection_About
Menu, Tray, Add, Quit, MenuSelection_Quit

Welcome_Message := "Hi. It looks like you're running the TEL Info Tool for the first time.`nWould you like to import the 'Intel Tool List' Excel file and use that data? If not, I will use the latest data from 3/1/2022."

;OnMessage(0x200, "WM_MOUSEMOVE")	                ; show a tooltip when hovering over controls for explaining control purpose (use control name + _TT)
if (Open_Workbook())
{
   Import_Sheets()
   Close_Workbook()
}
else
{
   MsgBox, 16, Workbook Missing, Unable to open the 'INTEL Tool List' workbook. The tool will close.
   ExitApp
}
Make_List()                                         ; build the tool_arrays and list of tools
populate_ddl()
GuiControl, Enable, MainDDL     				    ; was disabled to prevent click before it builds the list
Clipboard := ""
Clipboard := ClipSaved
Gui, Show, Autosize
OnMessage(0x201, "WM_LBUTTONDOWN")
WinGetPos,        PGuiX, PguiY, PGuiW, PGuiH, ahk_id %PGuihwnd%
Winset, Region,   % "w" PGuiW " h" PGuiH " 0-0 R15-15", ahk_id %PGuihwnd%

; === END AUTO-EXECUTE ===
; ========================

; ==============
; === LABELS ===
CloseMap:
Loop, %LoopTimes%
        {
            Gdip_DisposeImage(pBitmap%A_Index%)
        }
    Gui, 2:Destroy
	Return

PMinimize:
    WinMinimize
    Return

; ==============
; ===HOTKEYS ===

ESC::    ; TGR401
PExit:
If WinExist("ahk_id " PGuihwnd)
xl.quit
exitapp
return

; ================
; === INCLUDES ===
; TGR401
#Include Functions.ahk
#Include Map.ahk
