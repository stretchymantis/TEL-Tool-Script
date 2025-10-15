header          := false, hdr := ""
saveClipboard   := ClipboardAll
file2open       := A_ScriptDir "\INTEL Tool List Rev2.xlsx" ; put here the path of your excel file
xl              := ComObjCreate("Excel.Application")
Wrkbk           := xl.Workbooks.Open(file2open) 	
oRange 			:= xl.sheets("Layout").UsedRange
Loop % xl.sheets("Layout").UsedRange.Columns.count {
    hdr .= hdr ? "|" a_index : a_index
}
xl.sheets("Layout").UsedRange.copy
ClipWait, 1
data            := SubStr(clipboard,1,-2)
clipboard       := saveClipboard
Wrkbk.Close(0)
Gui,-DPIScale
loop,parse,data, `r,`n  
{
    If (A_Index = 1 ) {
        If (header)
            Gui, Add, ListView, x10 y10 vLV1 hdr grid hwndLV, % RegExReplace(a_loopfield, "`t", "|")
        else {
            Gui, Add, ListView, x10 y10 R55 vLV1 hdr grid hwndLV, % hdr
            LV_Add("", StrSplit(a_loopfield, a_tab)*)
        }
    }
    else			
        LV_Add("", StrSplit(a_loopfield, a_tab)*)
}
LV_ModifyCol(,"AutoHdr")
Gui, +Resize
VarSetCapacity(RECT, 16, 0)
SendMessage, LVM_GETITEMRECT := 0x100E, 0, &RECT, , ahk_id %LV%
GuiControl, Move, LV1, % "w"  NumGet(RECT, 8, "Int")+22*(A_ScreenDPI/96) ; SM_CXBORDER (1*2) + SM_CXVSCROLL (17) + SM_CXFIXEDFRAME(3)
Gui,show,AutoSize
return

GuiEscape:
GuiClose:
Esc::
  ExitApp
return