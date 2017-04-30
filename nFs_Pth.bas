Attribute VB_Name = "nFs_Pth"
Option Compare Database
Option Explicit

Sub PthAsst(Pth$, Optional Msg$)
If LasChr(Pth) <> "\" Then Er "Given {Pth} is not eof [\]", Pth
End Sub

Sub PthAsstExist(Pth$)
PthAsst Pth
If Not PthIsExist(Pth) Then Er "Given {Pth} does not exist", Pth
End Sub

Sub PthBrw(Pth)
Dim S$: S = FmtQQ("explorer ""?""", Pth)
Shell S, vbMaximizedFocus
End Sub

Sub PthClr(Pth$)
CurPthPush Pth
Dim I
On Error Resume Next
For Each I In PthFnAy(Pth)
    Kill I
Next
CurPthPop
End Sub

Function PthCpyFilUp1Dir(pDir$, Optional pFspc$ = "*.*", Optional ToPfx$ = "") As Boolean
'Aim: Copy all files of {pFspc} in {pDir} up 1 directory with a prefix {ToPfx}.  Target file will be overwritten.
Const cSub$ = "PthCpyFilUp1Dir"
Dim mChk As Boolean
mChk = False
''==Start
If Not IsDir(pDir) Then ss.A 1: GoTo E

'Copy Files in {pDir} up 1 directory
On Error GoTo R
Dim iFn$: iFn = VBA.Dir(pDir & pFspc)
Dim mA$
While iFn <> ""
    If Cpy_Fil(pDir & iFn, pDir & "..\" & ToPfx & iFn) Then mA = Add_Str(mA, iFn)
    iFn = VBA.Dir
Wend
If Len(mA) <> 0 Then ss.A 1, "Some files cannot be copied", eRunTimErr, "The Files", mA: GoTo E
If mChk Then
    MsgBox Fmt_Str("Check if all the of spec [{0}] the dir [{1}] is copied up 1 dir with pfx[{2}]", pFspc, pDir, ToPfx), vbInformation, "CopyFilUp1Dir"
    Opn_Dir pDir
    Stop
End If
Exit Function
R: ss.R
E:
End Function

Sub PthDltFil(Pth$, Optional FnSpec$ = "*.*")
If Not PthIsExist(Pth) Then Exit Sub
Dim FnAy$(): FnAy = PthFnAy(Pth, FnSpec)
If AyIsEmpty(FnAy) Then Exit Sub
Dim I
For Each I In FnAy
    FfnDlt Pth & I
Next
End Sub

Sub PthDltFnAy(Pth$, FnAy$())
Dim I
For Each I In FnAy
    FfnDlt Pth & I
Next
End Sub

Sub PthDltFnBySfx(Pth$, Sfx)
Dim F$(): F = PthFnAyBySfx(Pth, Sfx)
PthDltFnAy Pth, F
End Sub

Sub PthEns(Pth$)
If Not PthIsExist(Pth) Then MkDir Pth
End Sub

Sub PthEnsAllSeg(Pth$)
Dim A$(): A = Split(Pth, "\")
Dim P$, J%
P = A(0)
For J = 1 To UB(A)
    If A(J) <> "" Then
        P = P & "\" & A(J)
        PthEns P
    End If
Next
End Sub

Sub PthEnsAllSeg__Tst()
PthEnsAllSeg "C:\temp\a\a\a"
PthEnsAllSeg "C:\temp\a\a\a\"
RmDir "C:\temp\a\a\a"
RmDir "C:\temp\a\a"
RmDir "C:\temp\a"
End Sub

Function PthFnAy(Pth$, Optional FSpec$ = "*.*", Optional Atr As VbFileAttribute = vbNormal) As String()
PthAsst Pth
Dim A$
A = Dir(Pth & FSpec, Atr)
Dim O$()
While A <> ""
    Push O, A
    A = Dir
Wend
PthFnAy = O
End Function

Sub PthFnAy__Tst()
AyDmp PthFnAy("C:\Tmp\")
End Sub

Function PthFnAyBySfx(Pth$, Sfx, Optional Atr As VbFileAttribute = vbNormal) As String()
PthFnAyBySfx = AySel(PthFnAy(Pth, , Atr), "FfnHasSfx", Sfx)
End Function

Function PthFnnAy(Pth$, Optional FSpec$ = "*.*") As String()
Dim O$()
PthFnnAy = AyMapInto(PthFnAy(Pth, FSpec), O, "FfnCutExt")
End Function

Function PthIsExist(Pth$) As Boolean
PthIsExist = Fso.FolderExists(Pth)
End Function

Function PthNrm$(Pth$)
PthAsst Pth, "PthNrm"
Dim A$(): A = Split(Pth, "\")
Dim O$(), J%
J = UB(A)
Do While J >= 0
    If A(J) = ".." Then
        J = J - 2
    Else
        Push O, A(J)
        J = J - 1
    End If
Loop
O = AyRev(O)
PthNrm = Join(O, "\")
End Function

Sub PthNrm__Tst()
Dim P$: P = "c:\aa\..\"
Debug.Assert PthNrm(P) = "c:\"
End Sub
