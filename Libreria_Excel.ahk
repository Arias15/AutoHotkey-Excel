/*#SingleInstance,, Force
Browser_Forward::Reload
Browser_Back::
 
try XL := ComObjActive("Excel.Application") ;handle to running application
Catch {
    MsgBox % "no existing Excl ojbect:  Need to create one"
XL := ComObjCreate("Excel.Application")
XL.Visible := 1 ;1=Visible/Default 0=hidden
}
XL.Visible := 1 ;1=Visible/Default 0=hidden
MsgBox % "is an object? " IsObject(XL)
*/


1=Application 2=Workbook 3=Worksheet
XL_Handle(XL_Applic,1)    ;Application 
XL_Handle(XL_workbook,2)  ;Workbook
XL_Handle(XL_Worksheet,3) ;Worksheet
 
MsgBox % XL_Applic.name "`r"
 .   XL_workbook.name "`r"
 .   XL_Worksheet.name
 
 
XL_Handle(ByRef PXL,Sel){
ControlGet, hwnd, hwnd, , Excel71, ahk_class XLMAIN ;identify the hwnd for Excel
IfEqual,Sel,1, Return, PXL:= ObjectFromWindow(hwnd,-16).application ;Handle to Excel Application
IfEqual,Sel,2, Return, PXL:= ObjectFromWindow(hwnd,-16).parent ;Handlle to active Workbook
IfEqual,Sel,3, Return, PXL:= ObjectFromWindow(hwnd,-16).activesheet ;Handle to Active Worksheet
}
;***********adapted from ACC.ahk*******************
ObjectFromWindow(hWnd, idObject = -4){
(if Not h)?h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
 If DllCall("oleacc\AccessibleObjectFromWindow","Ptr",hWnd,"UInt",idObject&=0xFFFFFFFF,"Ptr",-VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
 Return ComObjEnwrap(9,pacc,1)
}

XL := ComObjActive("Excel.Application")
;***********************Show name of object handle is referencing********************************.
;XL_Reference(PXL) ;will pop up with a message box showing what pointer is referencing
XL_Reference(PXL){
 MsgBox, %HWND%
 MsgBox, % ComObjType(window)
 MsgBox % ComObjType(PXL,"Name")
}

;;********************Reference Cell by row and column number***********************************
XL.Range(XL.Cells(1,1).Address,XL.Cells(5,5).Address)
MsgBox % t 
MsgBox % XL.Cells(1,4).Value  ;Row, then column
XL.Range(XL.Cells(1,1).Address,XL.Cells(5,5).Address)
 
;***********************Screen update toggle********************************.
;XL_Screen_Update(XL)
XL_Screen_Update(PXL){
 PXL.Application.ScreenUpdating := ! PXL.Application.ScreenUpdating ;toggle update
}
;~ XL_Screen_Visibility(XL)
XL_Screen_Visibility(PXL){
 PXL.Visible:= ! PXL.Visible ;Toggle screen visibility
}
;***********************First row********************************.
XL_First_Row(XL)
XL_First_Row(PXL){
 Return, PXL.Application.ActiveSheet.UsedRange.Rows(1).Row
}

;***********************Used Rows********************************.
;~ Rows:=XL_Used_Rows(XL)
XL_Used_rows(PXL){
;  To do
}
;***********************Last Row********************************.
;~ LR:=XL_Last_Row(XL)
XL_Last_Row(PXL){
 Return PXL.Application.ActiveSheet.UsedRange.Rows(PXL.Application.ActiveSheet.UsedRange.Rows.Count).Row
}
;***********************First Column********************************.
;~ XL_First_Col_Nmb(XL)
XL_First_Col_Nmb(PXL){
 Return, PXL.Application.ActiveSheet.UsedRange.Columns(1).Column
}
;***********************First Column Alpha**********************************.
;~ XL_Last_Col_Alpha(XL)
XL_First_Col_Alpha(PXL){
 FirstCol:=PXL.Application.ActiveSheet.UsedRange.Columns(1).Column
 IfLessOrEqual,LastCol,26, Return, (Chr(64+FirstCol))
 Else IfGreater,LastCol,26, return, Chr((FirstCol-1)/26+64) . Chr(mod((FirstCol- 1),26)+65)
}
;***********************Used Columns********************************.
LC:=XL_Used_Cols_Nmb(XL)
XL_Used_Cols_Nmb(PXL){
    MsgBox, % PXL.Application.ActiveSheet.UsedRange.Columns.Count
 Return, PXL.Application.ActiveSheet.UsedRange.Columns.Count
}

;***********************Last Column********************************.
LC:=XL_Last_Col_Nmb(XL)
MsgBox NRColumnas %LC%
XL_Last_Col_Nmb(PXL){
 Return, PXL.Application.ActiveSheet.UsedRange.Columns(PXL.Application.ActiveSheet.UsedRange.Columns.Count).Column
}
;***********************Last Column Alpha**  Needs Workbook********************************.
;~ XL_Last_Col_Alpha(XL)
XL_Last_Col_Alpha(PXL){
 LastCol:=XL_Last_Col_Nmb(PXL)
 IfLessOrEqual,LastCol,26, Return, (Chr(64+LastCol))
 Else IfGreater,LastCol,26, return, Chr((LastCol-1)/26+64) . Chr(mod((LastCol- 1),26)+65)
}
;***********************Used_Range Used range********************************.
RG:=XL_Used_RG(XL,Header:=1) ;Use header to include/skip first row
MsgBox RANGO %RG%
XL_Used_RG(PXL,Header=1){
 IfEqual,Header,0,Return, XL_First_Col_Alpha(PXL) . XL_First_Row(PXL) ":" XL_Last_Col_Alpha(PXL) . XL_Last_Row(PXL)
 IfEqual,Header,1,Return, XL_First_Col_Alpha(PXL) . XL_First_Row(PXL)+1 ":" XL_Last_Col_Alpha(PXL) . XL_Last_Row(PXL)
}


;***********************Numeric Column to string********************************.
StrC := XL_Col_To_Char(26)
MsgBox NToS_Column %StrC%
XL_Col_To_Char(index){ ;Converting Columns to Numeric for Excel
 IfLessOrEqual,index,26, Return, (Chr(64+index))
 Else IfGreater,index,26, return, Chr((index-1)/26+64) . Chr(mod((index - 1),26)+65)
}
;***********************alpha to Number********************************.
StrToNum := XL_String_To_Number("ab")
MsgBox StrToNum %StrToNum%
XL_String_To_Number(Column){
 StringUpper, Column, Column
 Index := 0
 Loop, Parse, Column  ;loop for each character
 {ascii := asc(A_LoopField)
     if (ascii >= 65 && ascii <= 90)
 index := index * 26 + ascii - 65 + 1    ;Base = 26 (26 letters)
 else { return
 } }
return, index
}


;***********************Freeze Panes********************************.
;~ XL_Freeze(XL,Row:="1",Col:="B") ;Col A will not include cols which is default so leave out if unwanted
;***********************Freeze Panes in Excel********************************.
/*XL_Freeze(XL,Row:="1",Col:="B")
XL_Freeze(PXL,Row="",Col="A"){
 PXL.Application.ActiveWindow.FreezePanes := False ;unfreeze in case already frozen
 IfEqual,row,,return ;if no row value passed row;  turn off freeze panes
 PXL.Application.ActiveSheet.Range(Col . Row+1).Select ;Helps it work more intuitivly so 1 includes 1 not start at zero
 PXL.Application.ActiveWindow.FreezePanes := True
}
*/



;*******************************************************.
;***********************Formatting********************************.
;*******************************************************.
;***********************Alignment********************************.
XL_Format_HAlign(XL,RG:="A1:A10",h:=2) ;1=Left 2=Center 3=Right
XL_Format_HAlign(PXL,RG="",h="1"){ ;defaults are Right bottom
 IfEqual,h,1,Return,PXL.Application.ActiveSheet.Range(RG).HorizontalAlignment:=-4131 ;Left
 IfEqual,h,2,Return,PXL.Application.ActiveSheet.Range(RG).HorizontalAlignment:=-4108 ;Center
 IfEqual,h,3,Return,PXL.Application.ActiveSheet.Range(RG).HorizontalAlignment:=-4152 ;Right
}
MsgBox primerform
XL_Format_VAlign(XL,RG:="A1:A11",v:=4) ;1=Top 2=Center 3=Distrib 4=Bottom
XL_Format_VAlign(PXL,RG="",v="1"){
 IfEqual,v,1,Return,PXL.Application.ActiveSheet.Range(RG).VerticalAlignment:=-4160 ;Top
 IfEqual,v,2,Return,PXL.Application.ActiveSheet.Range(RG).VerticalAlignment:=-4108 ;Center
 IfEqual,v,3,Return,PXL.Application.ActiveSheet.Range(RG).VerticalAlignment:=-4117 ;Distributed
 IfEqual,v,4,Return,PXL.Application.ActiveSheet.Range(RG).VerticalAlignment:=-4107 ;Bottom
}
MsgBox SecondForm
;***********************Wrap text********************************.
;~ XL_Format_Wrap(XL,RG:="A1:B4",Wrap:=0) ;1=Wrap text, 0=no
XL_Format_Wrap(PXL,RG="",Wrap="1"){ ;defaults to Wrapping
 PXL.Application.ActiveSheet.Range(RG).WrapText:=Wrap
}
;***********Shrink to fit*******************
;~ XL_Format_Shrink_to_Fit(XL,RG:="A1",Shrink:=0) ;1=Wrap text, 0=no
XL_Format_Shrink_to_Fit(PXL,RG="",Shrink="1"){ ;defaults to Shrink to fit
 (Shrink=1)?(PXL.Application.ActiveSheet.Range(RG).WrapText:=0) ;if setting Shrink to fit need to turn-off Wrapping
 PXL.Application.ActiveSheet.Range(RG).ShrinkToFit :=Shrink
}



;***********************Merge / Unmerge cells********************************.
XL_Merge_Cells(XL,RG:="A12:B13",Warn:=0,Merge:=1) ;set to true if you want them merged
XL_Merge_Cells(PXL,RG,warn=0,Merge=0){ ;default is unmerge and warn off
 PXL.Application.DisplayAlerts := warn ;Warn about unmerge keeping only one cell
 PXL.Application.ActiveSheet.Range(RG).MergeCells:=Merge ;set merge for range
 (warn=0)?(PXL.Application.DisplayAlerts:=1) ;if warnings were turned off, turn back on
}

MsgBox CombinarCelda
;***********************Font size, type, ********************************.
XL_Format_Font(XL,RG:="A1:B1",Font:="Arial Narrow",Size:=25) ;Arial, Arial Narrow, Calibri,Book Antiqua
XL_Format_Font(PXL,RG="",Font="Arial",Size="11"){
 PXL.Application.ActiveSheet.Range(RG).Font.Name:=Font
 PXL.Application.ActiveSheet.Range(RG).Font.Size:=Size
}
MsgBox Tipo_Letra_Tamaño
;***********************Font bold, normal, italic, Underline********************************.
XL_Format_Format(XL,RG:="A1:B1",1) ; Bold:=1,Italic:=0,Underline:=3  Underline 1 thru 5
XL_Format_Format(PXL,RG="",Bold=0,Italic=0,Underline=0){
 PXL.Application.ActiveSheet.Range(RG).Font.Bold:= bold
 PXL.Application.ActiveSheet.Range(RG).Font.Italic:=Italic
 (Underline="0")?(PXL.Application.ActiveSheet.Range(RG).Font.Underline:=-4142):(PXL.Application.ActiveSheet.Range(RG).Font.Underline:=Underline+1)
}

MsgBox FormatoCell
;***********Cell Shading*******************
;2=none 3=Red 4=Lt Grn 5=Blue 6=Brt Yel 7=Mag 8=brt blu 15=Grey 17=Lt purp  19=Lt Yell 20=Lt blu 22=Salm 26=Brt Pnk
XL_Format_Cell_Shading(XL,RG:="A1:H1",Color:=28)
XL_Format_Cell_Shading(PXL,RG="",Color=0){
 PXL.Application.ActiveSheet.Range(RG).Interior.ColorIndex :=Color
}
MsgBox Formato_Sombreado

;***********************Cell Number format********************************.
XL_Format_Number(XL,RG:="A1:B4",Format:="#,##0") ;#,##0 ;0,000 ;0,00.0 ;0000 ;000.0 ;.0% ;$0 ;m/dd/yy ;m/dd ;dd/mm/yyyy
XL_Format_Number(PXL,RG="",format="#,##0"){
 PXL.Application.ActiveSheet.Range(RG).NumberFormat := Format
}
MsgBox Formato_Numero


;***********tab/Worksheet color*******************
;1=Black 2=White  3=Red 4=Lt Grn 5=Blue 6=Brt Yel 7=Mag 8=brt blu 15=Grey 17=Lt purp  19=Lt Yell 20=Lt blu 22=Salm 26=Brt Pnk
XL_Tab_Color(xl,"Hoja1","4")
XL_Tab_Color(PXL,Sheet_Name,Color){
 PXL.Sheets(Sheet_Name).Tab.ColorIndex:=Color ;color tab yellow
}
MsgBox Color_Sheet


;********************Select / Activate sheet***********************************
XL_Select_Sheet(XL,"Hoja2")
XL_Select_Sheet(PXL,Sheet_Name){
 PXL.Sheets(Sheet_Name).Select
}

MsgBox Selecionar_Hoja
;***********************Search- find text- Cell shading and Font color********************************.
;~ XL_Color(PXL:=XL,RG:="A1:D50",Value:="Joe",Color:="2",Font:=1) ;change the font color
XL_Color(PXL:=XL,RG:="A1:D50",Value:="Joe",Color:="1") ;change the interior shading
;***********************to do ********************************.
;*this is one or the other-  redo it so it does both***************.
XL_Color(PXL="",RG="",Value="",Color="1",Font="0"){
 if  f:=PXL.Application.ActiveSheet.Range[RG].Find[Value]{ ; if the text can be found in the Range
 first :=f.Address  ; save the address of the first found match
 Loop
 If (Font=0){
 f.Interior.ColorIndex :=Color, f :=PXL.Application.ActiveSheet.Range[RG].FindNext[f] ;color Interior & move to next found cell
 }Else{
 f.Font.ColorIndex :=Color, f :=PXL.Application.ActiveSheet.Range[RG].FindNext[f] ;color font & move to next found cell
 }Until (f.Address = first) ; stop looking when we're back to the first found cell
}}

MsgBox colorform

