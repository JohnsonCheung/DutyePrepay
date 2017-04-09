Attribute VB_Name = "nIde_nTth_Tth"
Option Compare Database
Option Explicit

Sub AA1()
TthLines__Tst
End Sub

Sub TthBrw_Md(Optional A As CodeModule)
AyBrw TthNy_Md(A)
End Sub

Sub TthBrw_Pj(Optional A As vbproject)
DrAyBrw AyStrBrk(TthNy_Pj(A), ".")
End Sub

Sub TthBrw_Pj__Tst()
'1 Declare
Dim A As vbproject

'2 Assign
A = 1

'3 Calling
TthBrw_Pj A

End Sub

Sub TthEns(Optional MthNm$, Optional A As CodeModule)
Dim OMd As CodeModule:
Dim OMthNm$
    Set OMd = MdNz(A)
    OMthNm = MthNmNz(MthNm, OMd)
If NmIsTstNm(OMthNm) Then Exit Sub

Dim OTstNm$
Dim OTthIsExist As Boolean
    OTstNm = NmToTstNm(OMthNm)
    OTthIsExist = MthIsInMd(OTstNm, OMd)
'---
If Not OTthIsExist Then
    Dim Lines
    Lines = TthLines(OMthNm, OMd)
    MdApdLines Lines, OMd
End If
MthShw_ForEdt OTstNm, OMd
End Sub

Sub TthEns__Tst()
TthEns "TthBrw_Pj", Md("nIde_nTth_Tth")
End Sub

Function TthFstBIdx&(Optional A As CodeModule)
Dim B As CodeModule: Set B = MdNz(A)
Dim C$(): C = MdBdyLy(A)
Dim J&
For J = 0 To UB(C)
    If SrcLinIsTth(C(J)) Then
        TthFstBIdx = J + B.CountOfDeclarationLines: Exit Function
    End If
Next
TthFstBIdx = -1
End Function

Function TthFstStru(Optional A As CodeModule) As MthStru
Dim B As CodeModule: Set B = MdNz(A)
Dim C&: C = TthFstBIdx(B): If C = -1 Then Exit Function
'TthFstStru = MdLyMthStru(MdLy(B), C)
End Function

Function TthIsInAnyMd(Optional A As CodeModule) As Boolean
Dim B$(): B = MdBdyLy(A)
If AyIsEmpty(B) Then Exit Function
Dim I, C As MthBrk
For Each I In B
    If SrcLinIsMth(I) Then
        C = MthBrkNew(I)
        If MthBrkIsTth(C) Then TthIsInAnyMd = True: Exit Function
    End If
Next
End Function

Function TthIsInAnyMd_Pri(Optional A As CodeModule) As Boolean
Dim B$(): B = MdBdyLy(A)
If AyIsEmpty(B) Then Exit Function
Dim I, C As MthBrk
For Each I In B
    If SrcLinIsMth(I) Then
        C = MthBrkNew(I)
        If MthBrkIsTth_Pri(C) Then TthIsInAnyMd_Pri = True: Exit Function
    End If
Next
End Function

Function TthIsInAnyMd_Pub(Optional A As CodeModule) As Boolean
Dim B$(): B = MdBdyLy(A)
If AyIsEmpty(B) Then Exit Function
Dim I, C As MthBrk
For Each I In B
    If SrcLinIsMth(I) Then
        C = MthBrkNew(I)
        If MthBrkIsTth_Pub(C) Then TthIsInAnyMd_Pub = True: Exit Function
    End If
Next
End Function

Function TthLines$(MthNm$, Optional A As CodeModule)
Dim J%
'----
Dim CMd As CodeModule
Dim CBrk As MthBrk
    Set CMd = MdNz(A)
    CBrk = MthBrkNewMthNm(MthNm, CMd)
    If MthBrkIsEmpty(CBrk) Then Er "TthLines: Given {MthNm} cannot in {Md}", MthNm, MdNm(CMd)
''-----------------------
Dim BPrmAy$()
Dim BPrmUB%
Dim BBrk As MthBrk
    BPrmAy = AyTrim(Split(CBrk.PrmStr, ","))
    BPrmUB = UB(BPrmAy)
    For J = 0 To BPrmUB
        BPrmAy(J) = RmvPfx(BPrmAy(J), "Optional ")
    Next
'
    'Remove BPrmAy Last Ele's Pfx-"Paramarray"
    If Sz(BPrmAy) >= 1 Then
        Dim LasPrm$: LasPrm = Pop(BPrmAy)
        Dim LasPrm1$: LasPrm1 = RmvPfx(LasPrm, "Paramarray ")
        Push BPrmAy, LasPrm1
    End If
'
    BBrk = CBrk
'''-------------
Dim APrmNmAy$()
Dim APrmSfxAy$()
Dim ARetSfx$
Dim ACallPrm$:
Dim AIsFct As Boolean:
Dim AIsRetObj As Boolean:
    Dim PrmBrkDrAy(): PrmBrkDrAy = AyMapInto(BPrmAy, PrmBrkDrAy, "PrmBrk")
    DrAyAsg PrmBrkDrAy, APrmNmAy, APrmSfxAy
    ARetSfx = BBrk.RetTyChr & BBrk.RetAs
    ACallPrm = Join(APrmNmAy, ", ")
    AIsFct = BBrk.Ty = "Function"
    AIsRetObj = MthBrkIsRetObj(BBrk)
'---------------------
Dim LSub$
Dim LDcl$()
Dim LAsg$()
Dim LCalling$()
Dim LAsst$()
Dim LEnd$
    LSub = FmtQQ("Sub ?()", NmToTstNm(MthNm))
    '
    Erase LDcl
        Push LDcl, "'1 Declare"
        For J = 0 To BPrmUB
            Push LDcl, "Dim " & BPrmAy(J)
        Next
        If AIsFct Then
            Push LDcl, "Dim Act" & ARetSfx
            Push LDcl, "Dim Exp" & ARetSfx
        End If
        Push LDcl, ""
    '
    Erase LAsg
        Push LAsg, "'2 Assign"
        For J = 0 To BPrmUB
            Push LAsg, APrmNmAy(J) & " = 1"
        Next
        If AIsFct Then
            Push LAsg, "Exp = 1"
        End If
        Push LAsg, ""
        
    Erase LCalling
        Push LCalling, "'3 Calling"
    
        Dim L$
        If AIsFct Then
            L = FmtQQ("Act = ?(?)", MthNm, ACallPrm)
            If AIsRetObj Then L = "Set " & L
        Else
            L = FmtQQ("? ?", MthNm, ACallPrm)
        End If
        Push LCalling, L
        If AIsFct Then
            Push LCalling, "Exp = 1"
        End If
        Push LCalling, ""
        
    Erase LAsst
        If AIsFct Then
            Push LAsst, "'4 Asst"
            Push LAsst, "Debug.Assert Act = Exp"
        End If
    LEnd = "End Sub"
'----------------------
TthLines = LyJn(ApSy(LSub, LDcl, LAsg, LCalling, LAsst, LEnd))
End Function

Sub TthLines__Tst()
Dim A$: A = TthLines("TthLines", Md("nIde_nTth_Tth"))
Debug.Print A
Stop
End Sub

Sub TthMov_Md(Optional A As CodeModule)
'Move all test method in Module-{A} to its Tst-Module.
Dim Md As CodeModule: Set Md = MdNz(A)          ' Md = A module    ! The codemodule of given module-A
Dim Tm As CodeModule: Set Tm = TmEns(Md)        ' Tm = Tst Module  ! Tst module of given module-A
Dim N%
Do
    N = N + 1: If N > 1000 Then Er "Impossible have >1000 Tst Mth in an module"
    Dim M As MthStru: M = TthFstStru(Md)      ' M  = Method Struc    ! first test method structure of module-A
                          If MthStruIsEmpty(M) Then Exit Sub
    Dim BEIdx&(): BEIdx = M.BEIdx
    Dim Lin$:       Lin = MdLinesByBEIdx(BEIdx, Md) ' Lin = Method Lines  ! Method lines of first test method
                          MthApd Lin, Tm
                          MdDltLin BEIdx, Md
Loop
End Sub

Sub TthMov_Pj(Optional A As vbproject)
End Sub

Sub TthRen_MdTo2DashSfxTst(Optional A As CodeModule)
TthRen_MdXXX A, "NmTo2DashSfxTst"
End Sub

Sub TthRen_MdToPfxTst2Dash(Optional A As CodeModule)
TthRen_MdXXX A, "NmToPfxTst2Dash"
End Sub

Sub TthRen_MdXXX(A As CodeModule, Fct$)
Const Dbg As Boolean = False
Dim OIdx_Lin_DrAy()
Dim OMd As CodeModule: Set OMd = MdNz(A)
    Dim OLy$()
    Erase OIdx_Lin_DrAy
    Dim TthNy$()
        OLy = MdLy(OMd)
        TthNy = TthNy_Md(OMd)
    Dim J%, NewM$, OldM$, NewLin$, Idx&, Dr
    Dim I&, L$
    For J = 0 To UB(TthNy)
        For I = 0 To UB(OLy)
            L = OLy(I)
            If InStr(L, TthNy(J)) = 0 Then GoTo NxtLin
            OldM = TthNy(J)
            NewM = Run(Fct, OldM)
            If NewM <> OldM Then
                NewLin = Replace(OLy(I), OldM, NewM)
                Idx = I
                Dr = Array(Idx, NewLin, OldM, NewM)
                Push OIdx_Lin_DrAy, Dr
            End If
NxtLin:
        Next
    Next
'----------
If Dbg Then Debug.Print AlignL(MdNm(OMd), 30);
If AyIsEmpty(OIdx_Lin_DrAy) Then
    If Dbg Then Debug.Print "<== No Tst-Mth":
    Exit Sub
End If
'DrAyBrw DrAyAddCol_ConstAtBeg(OIdx_Lin_DrAy, MdNm(OMd)): Stop
If Dbg Then Debug.Print "*** Some Method renamed ***"
For Each Dr In OIdx_Lin_DrAy
    AyAsg Dr, Idx, NewLin
    If Dbg Then Debug.Print vbTab; OLy(Idx)
    If Dbg Then Debug.Print vbTab; NewLin
    If Dbg Then Debug.Print
    Debug.Print NewLin
    MdRplLin Idx, NewLin, OMd  '<===
Next
MdSav OMd '<===
End Sub

Sub TthRen_PjTo2DashSfxTst(Optional A As vbproject)
TthRen_PjXXX A, "NmTo2DashSfxTst"
End Sub

Sub TthRen_PjTo2DashSfxTst__Tst()
TthRen_PjTo2DashSfxTst
End Sub

Sub TthRen_PjToPfxTst2Dash(Optional A As vbproject)
TthRen_PjXXX A, "NmToPfxTst2Dash"
End Sub

Sub TthRen_PjXXX(A As CodeModule, Fct$)
Dim MdAy() As CodeModule: MdAy = PjMdAy(A)
AyEachEle MdAy, "TthRen_MdXXX", Fct
End Sub

Sub TthRun(Optional A As CodeModule)
Dim ATth$(): ATth = TthNy_Md(A)
'------------------
Dim O$()
    O = AySel(ATth, "NmIsTstNm")
    O = AyMinus(O, Array("Tst__" & "TthRun", "TthRun" & "__Tst"))
    O = AyExcl(O, "StrHas", "TthRen")

'-------------------
If AyIsEmpty(O) Then Exit Sub
Dim Mth
For Each Mth In O
    Run Mth
Next
End Sub

Sub TthRun__Tst()
TthRun
End Sub
