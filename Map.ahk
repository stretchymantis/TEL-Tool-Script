Show_Map() {
   Global
   ClipSaved := ClipboardAll
   pi:=3.14159265
   ;gdip init
   pToken 		:= Gdip_Startup()                                     ; Import the Layout Map
   If !pToken  := Gdip_Startup() {
	   MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
	   ExitApp
   }
   SysGet, Mon, MonitorWorkArea                                   ; Get monitor resolution for fullscreen map (MonLeft, MonTop, MonRight, MonBottom)
   ; DELETE ; ToolTip, Left: %MonLeft%`n Right: %monright% `n Bottom: %monbottom%
   LoopTimes := 1
   open_workbook()
   xl.sheets("Layout").Activate
   xl.ActiveWindow.DisplayGridlines :=	False
   ;xl.sheets(wksht).UsedRange.copy
   xl.range("A1:AO59").CopyPicture(1,2)                              ; Copy Layout Map
      if !(pBitMapOriginal_Got = "1")
      {
      pBitMapOriginal_Got := 1
      pBitmapOriginal := Gdip_CreateBitmapFromClipboard()               ; Original to replace when calling another new map
      }
   pBitmap1 := pBitmapOriginal                 ; Import the Layout Map
   pBitmap2 := pBitmapOriginal                 ; Import the Layout Map
   if mainDDL                                                        ; If user selected a tool then clicked the map (as opposed to it being blank)
   {   
      FoundCell := xl.ActiveSheet.UsedRange.Find(mainDDL)            ; Find tool number from mainDDL
      if (FoundCell <> "")
      {
         rangeArray := [FoundCell.Address]                           ; List of cell blocks which should be colored (will be completed later)
         for i in rangeArray                                         ; Cycle through all elements of rangeArray 
         {
            currentRange := rangeArray[i]                            ; Access elements of rangeArray
            for index, Side in [7, 8, 9, 10]                         ; Left, Top, Bottom, Right 
            {
               ; --- Get original cell properties before changing them
               Cell_Orig_Border_Color     := xl.Range(currentRange).Borders(Side).ColorIndex
               Cell_Orig_Border_LineStyle := xl.Range(currentRange).Borders(Side).LineStyle
               Cell_Orig_Border_Weight    := xl.Range(currentRange).Borders(Side).Weight
            }
               Xl.Range(FoundCell.Address).Select
               Cell_Orig_Font_Style := xl.Selection.Font.Fontstyle 
               Cell_Orig_Font_Size  := xl.Selection.Font.Size
               Cell_Orig_Font_Color := xl.Selection.Font.ColorIndex

               ; --- Change selected cell properties for 'flash' state
               for index, Side in [7, 8, 9, 10] 
               {
               xl.Range(currentRange).Borders(Side).ColorIndex := 3
               ;xl.Range(currentRange).Borders(Side).LineStyle := 1   ; Continous line
               ;xl.Range(currentRange).Borders(Side).Weight := 3
               }
               Xl.Range(FoundCell.Address).Select                    ; FontStyle & ColorIndex only works when using 'Selection.Font.' for some reason
               ;Xl.Selection.Font.Fontstyle := "regular"              ; regular, italic, bold, bold italic
               ;Xl.Selection.Font.Size := 8
               Xl.Selection.Font.ColorIndex := 3                     ; 3 = Red
               Sleep 150                                             ; Throws an error here sometimes unless slowed down
               xl.range("A1:AO59").CopyPicture(1,2)                  ; Copy Layout Map to clipboard
               pBitmap2 := Gdip_CreateBitmapFromClipboard()          ; Import the Layout Map
               /*
               ; --- Revert cell properties back to original
               for index, Side in [7, 8, 9, 10]                         ; Left, Top, Bottom, Right 
               {
               xl.Range(currentRange).Borders(Side).ColorIndex    := Cell_Orig_Border_Color    
               xl.Range(currentRange).Borders(Side).LineStyle     := Cell_Orig_Border_LineStyle
               xl.Range(currentRange).Borders(Side).Weight        := Cell_Orig_Border_Weight 
               }  
               Xl.Range(FoundCell.Address).Select
               xl.Selection.Font.Fontstyle                        := Cell_Orig_Font_Style
               xl.Selection.Font.ColorIndex                       := Cell_Orig_Font_Color
               xl.Selection.Font.Size                        := Cell_Orig_Font_Size            
         */
         }
      }
   LoopTimes := 2
   }

Close_Workbook()
map_width := Gdip_GetImageWidth(pBitmap1), map_height := Gdip_GetImageHeight(pBitmap1) ; Get the width and height of the image to resize it if too big
orig_map_width := map_width
orig_map_height := map_height
Ratio := map_width/map_height

; resize map to fit screen
map_height := (A_Screenheight - (A_Screenheight * .25))
map_width := map_height * ratio
;   map_width := (A_Screenwidth - (A_Screenwidth * .5))
;   map_height := map_width*(1/Ratio)

SysGet, monitor_, MonitorWorkArea, 1
monitor_Width := monitor_Right - monitor_Left
monitor_Height := monitor_Bottom - monitor_Top
monitor_center_x := (monitor_Width - map_width)/2
monitor_center_y := (monitor_Height - map_height)/2

Gui, 2:  +LastFound E0x80000 +hwndMGuihwnd            ; create GUI for map
Gui, 2: Show, NA
winHwnd := WinExist()

hbm := CreateDIBSection(map_width, map_height)
hdc := CreateCompatibleDC()
;hdc2 := CreateCompatibleDC()
      obm := SelectObject(hdc, hbm)
      G := Gdip_GraphicsFromHDC(hdc)
Gdip_SetSmoothingMode(G, 4)
Gdip_SetInterpolationMode(G, 7)
Gdip_DrawImage(G, pBitmap1, 0, 0, map_width, map_height, 0, 0, orig_map_width, orig_map_height)
Gdip_DrawImage(G, pBitmap2, 0, 0, map_width, map_height, 0, 0, orig_map_width, orig_map_height)
Options = x-75 y-12 Right cff000000 r4 s27 bold
Gdip_TextToGraphics(G, "_", Options, "Arial", map_width)
Options = x-15 y5 Right cff000000 r4 s17
Gdip_TextToGraphics(G, "X", Options, "Arial", map_width)
pBrush := Gdip_BrushCreateSolid(0xffff0000)
pPen := Gdip_CreatePen(0xffff0000, 3)
UpdateLayeredWindow(winHwnd, hdc, monitor_center_x, monitor_center_y, map_width, map_height)
WinActivate
if LoopTimes = 2
   SetTimer, FlashCell, on
Clipboard := ClipSaved
}

FlashCell() {
Global
   loop, 2
   {
      Gdip_DrawImage(G, pBitmap%A_index%, 0, 0, map_width, map_height, 0, 0, orig_map_width, orig_map_height)
      Gdip_DrawImage(G, pBitmap%A_Index%, 0, 0, map_width, map_height, 0, 0, orig_map_width, orig_map_height)
      Options = x-75 y-12 Right cff000000 r4 s27 bold
      Gdip_TextToGraphics(G, "_", Options, "Arial", map_width)
      Options = x-15 y5 Right cff000000 r4 s17
      Gdip_TextToGraphics(G, "X", Options, "Arial", map_width)Gdip_TextToGraphics(G, "X", Options, "Arial", map_width)
      UpdateLayeredWindow(winHwnd, hdc,,, map_width, map_height)
      sleep 200
   }
}
;--------------------------------------------------------
/*
MoveControl:                                                ; Move map control inside stationary window
      MouseGetPos,,,,winhwnd, 2
      PostMessage, 0x112,0xF012,0,,ahk_id %winhwnd%         ; [ WM_SYSCOMMAND+SC_MOVE ]
      Winset,Redraw,,ahk_id %winhwnd%                       ; Thanks to adamrgolf
      KeyWait, LButton
      GuiGetPos( MGuiX,MGuiY,MGuiW,MGuiH, MGuihwnd )           ; More accurate for getting position; otherwise have to use offset -16 for x & -39 for y
      GuiControlGet, Img1, Pos, Img1
      Img1XEnd := Img1W + Img1X
      Img1YEnd := Img1H + Img1Y
      Img1XOverlapped := MGuiW - Img1W
      Img1YOverlapped := MGuiH - Img1H
      If (Img1XEnd < MGuiw)
      GuiControl, MoveDraw, Img1, % "X" Img1XOverlapped        ; Restrict map from moving its right edge left past right GUI border
      If (Img1YEnd < MGuiH)
      GuiControl, MoveDraw, Img1, % "Y" Img1YOverlapped        ; Restrict map from moving its bottom edge above the bottom GUI border
      If (Img1X > 0)
      GuiControl, MoveDraw, Img1, % "X" 0                      ; Restrict map from moving its left edge past left GUI border
      If (Img1Y > 0)                                           ; Restrict map from moving its top edge below top GUI border
      GuiControl, MoveDraw, Img1, % "Y" 0
      WinSet, Top,, ahk_id %ScaleCBhwnd%                       ; Bring 'Scale when sizing' checkbox back to top map
      old_zoom := 1.0
      Return
*/