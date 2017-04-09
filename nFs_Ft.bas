Attribute VB_Name = "nFs_Ft"
Option Compare Database
Option Explicit

Sub FtBrw(Ft$, Optional DltFt As Boolean, Optional WinSty As VbAppWinStyle = VbAppWinStyle.vbMaximizedFocus)
Const Pgm$ = "NotePad"
Dim S$: S = FmtQQ("""?"" ""?""", Pgm, Ft)
Shell S, WinSty
End Sub

Sub FtBrw__Tst()
Dim F$
F = TmpFt
AyWrt Split("a b cd"), F
FtBrw F
End Sub

Function FtCmp(oIsSam As Boolean, Ft1$, Ft2$) As Boolean
'Aim: compare if 2 files are the same.
Const cSub$ = "FtCmp"
Const cBlkSiz% = 8192
On Error GoTo R
oIsSam = False
If VBA.Dir(Ft1) = "" Then ss.A 1, "Ft1 not exist": GoTo E
If VBA.Dir(Ft2) = "" Then ss.A 2, "Ft2 not exist": GoTo E
Dim N1%: N1 = VBA.FileSystem.FileLen(Ft1)
Dim N2%: N2 = VBA.FileSystem.FileLen(Ft2)
If N1 <> N2 Then Exit Function
Dim mF1 As Byte, mF2 As Byte
mF1 = FreeFile: Open Ft1 For Binary Access Read As mF1
mF2 = FreeFile: Open Ft2 For Binary Access Read As mF2
Dim J%, NBlk%
NBlk% = ((N1 - 1) \ cBlkSiz) + 1
For J = 0 To NBlk% - 1
    Dim mA1$: mA1 = Input(cBlkSiz, mF1)
    Dim mA2$: mA2 = Input(cBlkSiz, mF2)
    If mA1 <> mA2 Then
        Close mF1, mF2
        Exit Function
    End If
Next
Close mF1, mF2
oIsSam = True
Exit Function
R: ss.R
E: FtCmp = True: ss.B cSub, cMod, "Ft1,Ft2", Ft1, Ft2
End Function

Function FtCmp__Tst()
Dim mF1 As Byte, mF2 As Byte
Const mFt1$ = "c:\a1.txt"
Const mFt2$ = "c:\a2.txt"
mF1 = FreeFile: Open mFt1 For Output Access Write As mF1
mF2 = FreeFile: Open mFt2 For Output Access Write As mF2
Dim J%, mB$: mB$ = "0123456789"
Dim mC$: mC = "0123456789"
For J = 1 To 900
    Print #mF1, mB
    If J = 900 Then
        Print #mF2, mC
    Else
        Print #mF2, mB
    End If
Next
Close mF1, mF2
Dim mIsSam As Boolean: If FtCmp(mIsSam, mFt1, mFt2) Then Stop
Debug.Print mIsSam
End Function

Function FtLines$(Ft)
FtLines = LyJn(FtLy(Ft))
End Function

Function FtLy(Ft) As String()
Dim F%: F = FtOpnInp(Ft)
Dim L$, O$()
While Not EOF(Ft)
    Line Input #F, L
    Push O, L
Wend
FtLy = O
Close (F)
End Function

Function FtOpnInp%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Input As #O
FtOpnInp = O
End Function

Function FtOpnOup%(Ft)
Dim O%: O = FreeFile(1)
Open Ft For Output As #O
FtOpnOup = O
End Function
