Attribute VB_Name = "nVb_nZip_Ffn"
Option Compare Database
Option Explicit
Const Pgm$ = "C:\Program Files\7-zip\7z.exe"

Sub FfnUnZip(Ffn)
FfnAsstExt Ffn, ".zip", "FfnUnzip"
Dim Fzip$
    Fzip = FfnRplExt(Ffn, ".zip")
FfnAsstNotExist Fzip, "FfnUnZip"
Dim oCmd$
oCmd = FmtQQ("""?"" x -p20071122 ""?"" ""?""", Pgm, Fzip, Ffn)
Shell oCmd, vbHide
End Sub

Sub FfnZip(Ffn)
If Right(Ffn, 4) = ".zip" Then Er "Given {Ffn} cannot end with .zip", Ffn
FfnAsstExist Ffn, "FfnZip"
Dim Fzip$
    Fzip = FfnRplExt(Ffn, ".zip")
Dim oCmd$
    oCmd = FmtQQ("""?"" a -p20071122 ""?"" ""?""", Pgm, Fzip, Ffn)
Shell oCmd, vbHide
End Sub

Sub FfnZip__Tst()
Dim mWb As Workbook: If Crt_Wb(mWb, "c:\aa.xls") Then Stop
If Cls_Wb(mWb, True) Then Stop
FfnZip "C:\aa.xls"
End Sub

Sub UnZip__Tst()
FfnUnZip "c:\aa.zip"
End Sub
