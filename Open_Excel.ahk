;***********************Open********************************.
;***********************open excel********************************.
XL_Open(XL,Vis:=1,Try:=1,Path:="F:\rosa\Desktop\Adberto\Programas\AutoHotkey\Excel\PruebaMacro.xlsm") ;XL is pointer to workbook, Vis=0 for hidden Try=0 for new Excel

XL_Open(PXL,vis=1,Try=1,Path=""){
 If (Try=1){
 Try PXL := ComObjActive("Excel.Application") ;handle
 Catch
 PXL := ComObjCreate("Excel.Application") ;handle
 PXL.Visible := vis ;1=Visible/Default 0=hidden
 }Else{
 PXL := ComObjCreate("Excel.Application") ;handle
 PXL.Visible := vis ;1=Visible/Default 0=hidden
 }
 PXL:=PXL.Workbooks.Open(path) ;wrb =handle to specific workbook
 Return,PXL
}
epExcel:= ComObjActive("Excel.Application")

epExcel.Run("Macro1")

;epExcel.Workbooks.Add()
epExcel.Sheets.Add()
;epExcel:=

MsgBox final
return