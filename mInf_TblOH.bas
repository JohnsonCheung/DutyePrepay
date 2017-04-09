Attribute VB_Name = "mInf_TblOH"
Option Compare Database
Option Explicit

Function TblOHMaxYMD() As YMD
Dim M$: M = 20000000 + SqlLng("Select Max(YY*10000+MM*100+DD) from OH")
TblOHMaxYMD = DteYMD(CDate(Format(M, "0000-00-00")))
End Function

Sub TblOHMaxYMD__Tst()
Debug.Print YMDToStr(TblOHMaxYMD)
End Sub
