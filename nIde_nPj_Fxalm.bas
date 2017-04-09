Attribute VB_Name = "nIde_nPj_Fxalm"
Option Compare Database
Option Explicit

Function FxlamCrt(Fxlam$) As vbproject
Dim X As Excel.Application
Dim W As Workbook
Set X = Appx
FfnAsstExt Fxlam, ".xlam", "FxlamCrt"
Set W = Appx.Workbooks.Add
VbeLasPj(Appx.Vbe).Name = FfnFnn(Fxlam)
WbSavAs W, Fxlam, xlOpenXMLAddIn
WbCls W, NoSav:=True
Set FxlamCrt = OpnPjFxlam(Fxlam)
End Function
