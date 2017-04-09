Attribute VB_Name = "nIde_Res"
Option Compare Database
Option Explicit
Const C_Mod$ = "nIde_Res"

Function ResDt(ResMthNm$, Optional ResMdNm$, Optional PjNm$) As Dt
ResDt = DtNewScLy(ResLy(ResMthNm, ResMdNm, PjNm))
End Function

Sub ResDt__Dt()
'Tbl;ABC
'Fld;AA;BB;CC
';1;2;3
';2;3;4
End Sub

Sub ResDt__Tst()
'1 Declare
Dim ResMthNm$
Dim ResMdNm$
Dim PjNm$
Dim Act As Dt
Dim Exp As Dt

'2 Assign
ResMthNm = "ResDt__Dt"
ResMdNm = C_Mod
PjNm = C_Lib
Exp = DtNewScLy(SplitVBar("Tbl;ABC|Fld;AA;BB;CC|;1;2;3|;2;3;4"))

'3 Calling
Act = ResDt(ResMthNm, ResMdNm, PjNm)

'4 Asst
DtAsstEq Act, Exp
End Sub

Function ResLy(ResMthNm$, Optional ResMdNm$, Optional PjNm$) As String()
Dim M As CodeModule
    Set M = Md(ResMdNm, Pj(PjNm))
ResLy = AyRmvFstChr(AyRmvLasEle(AyRmvAt(MthLy(ResMthNm, M), 0)))
End Function

Sub ResLy__Tst()
AyAsstEq ResLy("ResLyRes__Tst"), ApSy("", "aa", "", "bb")
End Sub

Function ResLyMd(ResMdNm$, Optional PjNm$) As String()
Dim M As CodeModule: M = Md(ResMdNm, Pj(PjNm))
ResLyMd = AyRmvFstChr(MdLy(M))
End Function

Sub ResLyRes__Tst()
'
'aa
'
'bb
End Sub

Function ResStr$(ResMthNm$, Optional ResMdNm$, Optional PjNm$)
ResStr = LyJn(ResLy(ResMthNm, ResMdNm, PjNm))
End Function
