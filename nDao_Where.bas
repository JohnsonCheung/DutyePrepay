Attribute VB_Name = "nDao_Where"
'Option Compare Text
'Option Explicit
'Option Base 0
Const cMod$ = ""

Function Bld_Struct_ForTy$(pItm$, Optional pN As Byte = 0, Optional pX$ = "")
'Aim: Build {mR} (a list of field (may with field type and len for table creation) from {pItm}, {pN}, {pX}
'           pN is 0 to 5 (MaxTy)
'           pX is "", "x", .., "xxx"
'       Ty Tables:   name = [$<pItm>Ty]     Example, $TblTy for each record in $Tbl.  3 fields: Tbl, TblTy1, TblTy2
'       Ty1 Tables:  name = [$<pItm>Ty1]    Example, $TblTy1.                         4 fields: TblTy1,    NmTblTy1,    TblTy1x,   DesTblTy1
'                           [$<pItm>Ty1x]   Example, $TblTy1x                         4 fields: TblTy1x,   NmTblTy1x,   TblTy1xx,  DesTblTy1
'                           [$<pItm>Ty1xx]  Example, $TblTy1xx                        4 fields: TblTy1xx,  NmTblTy1xx,  TblTy1xxx, DesTblTy1
'                           [$<pItm>Ty1xxx] Example, $TblTy1xxx                       3 fields: TblTy1xxx, NmTblTy1xxx,            DesTblTy1
'       Ty2 Tables:  name = [$<pItm>Ty2]    Example, $TblTy2.                         4 fields: TblTy2,    NmTblTy2,    TblTy2x,   DesTblTy2
'                           [$<pItm>Ty2x]   Example, $TblTy2x                         4 fields: TblTy2x,   NmTblTy2x,   TblTy2xx,  DesTblTy2
'                           [$<pItm>Ty2xx]  Example, $TblTy2xx                        4 fields: TblTy2xx,  NmTblTy2xx,  TblTy2xxx, DesTblTy2
'                           [$<pItm>Ty2xxx] Example, $TblTy2xxx                       3 fields: TblTy2xxx, NmTblTy2xxx,            DesTblTy2
'Note: pN=0, pX will be "1", .., "5" (which means MaxTy), Ty Table will be return
Const cSub$ = "Bld_Struct_ForTy"
On Error GoTo R
Dim mR$
If pN = 0 Then
    If 1 > Val(pX) Or Val(pX) > 5 Then ss.A 1, "pN=0, pX must be '1', ..,'5', which means MaxTy", ePrmErr: GoTo E
ElseIf pN > 5 Then
    ss.A 2, "pN should 0-5", ePrmErr: GoTo E
Else
    If pX <> "" And pX <> "x" And pX <> "xx" And pX <> "xxx" Then ss.A 1, "If pN between 1-5, pX must be '','x',..'xxx'", ePrmErr: GoTo E
End If
', Optional pForCrt As Boolean = False
'If pForCrt Then Exit Function
If pN = 0 Then
    Dim J%
    mR = pItm
    For J = 1 To CByte(pX)
        mR = mR & ",Ty" & pItm & J
    Next
    Exit Function
End If
If pX = "xxx" Then
    mR = Fmt_Str("Ty{0}{1}{2},NmTy{0}{1}{2},DesTy{0}{1}{2}", pItm, pN, pX)
Else
    mR = Fmt_Str("Ty{0}{1}{2},NmTy{0}{1}{2},Ty{0}{1}{2}x,DesTy{0}{1}{2}", pItm, pN, pX)
End If
Bld_Struct_ForTy = mR
Exit Function
R: ss.R
E:
End Function

Function Bld_Struct_ForTy__Tst()
Debug.Print Bld_Struct_ForTy_Import("Tbl", 3)
Debug.Print Bld_Struct_ForTy("Tbl", 0, "3")
Debug.Print Bld_Struct_ForTy("Tbl", 1, "")
Debug.Print Bld_Struct_ForTy("Tbl", 1, "x")
Debug.Print Bld_Struct_ForTy("Tbl", 1, "xx")
Debug.Print Bld_Struct_ForTy("Tbl", 1, "xxx")
Debug.Print Bld_Struct_ForTy("Tbl", 2, "")
Debug.Print Bld_Struct_ForTy("Tbl", 2, "x")
Debug.Print Bld_Struct_ForTy("Tbl", 2, "xx")
Debug.Print Bld_Struct_ForTy("Tbl", 2, "xxx")
Debug.Print Bld_Struct_ForTy("Tbl", 3, "")
Debug.Print Bld_Struct_ForTy("Tbl", 3, "x")
Debug.Print Bld_Struct_ForTy("Tbl", 3, "xx")
Debug.Print Bld_Struct_ForTy("Tbl", 3, "xxx")
Shw_DbgWin
End Function

Function Bld_Struct_ForTy_Import$(pItm$, pMaxTy As Byte)
'Aim: Build {mR} of a Import table [>{pItm}] having {pMaxTy}>0
'     Assume in there is import table named as [>{pItm}] and pMaxTy=2, the return {mR} will be
'       Nm<pItm>, and,
'       Nm<pItm>Ty1, Nm<pItm>Ty1x, Nm<pItm>Ty1xx, Nm<pItm>Ty1xxx, and,
'       Nm<pItm>Ty2, Nm<pItm>Ty2x, Nm<pItm>Ty2xx, Nm<pItm>Ty2xxx.
'     Eg <pItm>=Tbl
'       NmTbl, and,
'       NmTblTy1, NmTblTy1x, NmTblTy1xx, NmTblTy1xxx, and,
'       NmTblTy2, NmTblTy2x, NmTblTy2xx, NmTblTy2xxx.
Const cSub$ = "Bld_Struct_ForTy_Import"
Dim mR$
If pMaxTy = 0 Then ss.A 1, "pMaxTy must >0", , "pItm,pMaxTy", pItm, pMaxTy: GoTo E
On Error GoTo R
mR = "Nm" & pItm
Dim J%
For J = 1 To pMaxTy
    mR = mR & Fmt_Str(",NmTy{0}{1},NmTy{0}{1}x,NmTy{0}{1}xx,NmTy{0}{1}xxx", pItm, J)
Next
Bld_Struct_ForTy_Import = mR
Exit Function
R:
E:
End Function

Function Bld_Struct_ForTy_Import__Tst()
Debug.Print Bld_Struct_ForTy_Import("Tbl", 2)
End Function

Function WhereBld$(pLp$, pVayv)
'Aim:
Dim OWhere$
If pLp = "" Then Exit Function
Dim mAn$(): mAn = Split(pLp$, CtComma)
Dim mAv$(): mAv = pVayv
Dim N1%: N1 = Siz_Ay(mAn)
Dim N2%: N2 = Siz_Ay(mAv)
If N1 <> N2 Then Er "Count in pLp & pVayv mismatch", "Cnt in pLp, Cnt in pVayv", N1, N2
Dim J%: For J = 0 To N1 - 1
    Dim mA$: mA = mAn(J) & " in (" & mAv(J) & ")"
    OWhere = Add_Str(OWhere, mA, " and ")
Next
If OWhere <> "" Then OWhere = " Where " & OWhere
WhereBld = OWhere
End Function

Function WhereBld__Tst()
Dim mWhere$, mAv$(2)
mAv(0) = "11,22,33"
mAv(1) = "44,55,66"
mAv(2) = "77,88"
mWhere = WhereBld("aa,bb,cc", mAv)
Debug.Print mWhere
End Function
