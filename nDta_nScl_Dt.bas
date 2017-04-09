Attribute VB_Name = "nDta_nScl_Dt"
Option Compare Database
Option Explicit

Function DtNewSclVBar(SclVBar$) As Dt
DtNewSclVBar = DtNewScLy(SplitVBar(SclVBar))
End Function

Sub DtNewSclVBar__Tst()
DtBrw DtNewSclVBar("Tbl;ABC|Fld;A;B;C;D|;1;2;3;4|;4;4;5;1|;SLKF;DKF;SDFLDF;DFDF|;1")
End Sub

Function DtNewScLy(ScLy) As Dt
Dim L0$: L0 = ScLy(0)
Dim L1$: L1 = ScLy(1)

Dim Tn$
Dim Fny$():
Dim DrAy():
    Tn = RmvPfx(L0, "Tbl;")
    Fny = SclSy(RmvPfx(L1, "Fld;"))
    Dim B$(): B = AyRmvFstChr(AyRmvAt(ScLy, 0, 2))
    DrAy = DrAyNewScLy(B)
DtNewScLy = DtNew(Fny, DrAy, Tn)
End Function

Sub DtNewScLy__Tst()
DtBrw DtNewScLy(SplitVBar("Tbl;ABC|Fld;A;B;C;D|;1;2;3;4|;4;4;5;1|;SLKF;DKF;SDFLDF;DFDF|;1"))
End Sub

Function DtRead(Ft$) As Dt
DtRead = DtNewScLy(SyRead(Ft))
End Function

Sub DtRead__Tst()
Dim T$: T = TmpFt
DtWrt DtSample2, T
FtBrw T
Stop
Dim Act As Dt
Act = DtRead(T)
DtBrw Act
Stop
'DtAsstEq T, Act
End Sub

Function DtScLy(A As Dt) As String()
Dim O$()
Push O, "Tbl;" & A.Tn
Push O, "Fld;" & DrScl(A.Fny)
Dim J&
For J = 0 To UB(A.DrAy)
    Push O, ";" & DrScl(A.DrAy(J))
Next
DtScLy = O
End Function

Sub DtScLy__Tst()
AyBrw DtScLy(DtSample2)
End Sub

Sub DtWrt(A As Dt, Ft$)
AyWrt DtScLy(A), Ft
End Sub
