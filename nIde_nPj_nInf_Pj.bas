Attribute VB_Name = "nIde_nPj_nInf_Pj"
Option Compare Database
Option Explicit
Enum ePjOf
    ePjx
    ePja
    ePjo
    ePjp
    ePjw
End Enum

Function PjAltPjFfn$(Optional A As vbproject)
Dim Ap As vbproject: Set Ap = PjNz(A)
If PjIsTj(Ap) Then
    Dim P$: P = PthNrm(PjPth(Ap) & "..\")
    PjAltPjFfn = P & RmvPfx(Ap.Name, "Tst_") & AppExtNrm
Else
    PjAltPjFfn = TjFfn(Ap)
End If
End Function

Function PjAppa(A As vbproject) As Access.Application
If VbeHasPj(A, Application.Vbe) Then Set PjAppa = Application: Exit Function
If VbeHasPj(A, Appa.Vbe) Then Set PjAppa = Appa: Exit Function
End Function

Function PjCmpAy(Optional A As vbproject) As VBComponent()
Dim P As vbproject: Set P = PjNz(A)
Dim O() As VBComponent
Dim C As VBComponent
For Each C In P.VBComponents
    If C.Type = vbext_ct_StdModule Or C.Type = vbext_ct_ClassModule Then
        PushObj O, C
    End If
Next
PjCmpAy = O
End Function

Function PjCur() As vbproject
Set PjCur = Application.Vbe.ActiveVBProject
End Function

Function PjHasMdNm(MdNm$, Optional Pj As vbproject) As Boolean
Dim C As VBComponent
For Each C In PjNz(Pj).VBComponents
    If C.Name = MdNm Then PjHasMdNm = True: Exit Function
Next
End Function

Function PjIsOpn(PjNm$, Optional Vbe As Vbe) As Boolean
Dim P As vbproject
For Each P In VbeNz(Vbe).VBProjects
    If P.Name = PjNm Then PjIsOpn = True: Exit Function
Next
End Function

Function PjIsTj(Pj As vbproject) As Boolean
PjIsTj = IsPfx(Pj.Name, "Tst_")
End Function

Function PjMdAy(Optional A As vbproject, Optional MdNmLik$ = "*") As CodeModule()
Dim C As VBComponent
Dim M As CodeModule
Dim O() As CodeModule
For Each C In PjNz(A).VBComponents
    Select Case C.Type
    Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document
        Set M = C.CodeModule
        If Not MdIsEmpty(M) Then
            PushObj O, M
        End If
    End Select
Next
PjMdAy = O
End Function

Function PjMdAy_Cls(Optional A As vbproject, Optional MdNmLik$ = "*") As CodeModule()
Dim C As VBComponent
Dim M As CodeModule
Dim O() As CodeModule
For Each C In PjNz(A).VBComponents
    If C.Type = vbext_ct_ClassModule Then
        Set M = C.CodeModule
        If Not MdIsEmpty(M) Then
            PushObj O, M
        End If
    End If
Next
PjMdAy_Cls = O
End Function

Function PjMdAy_Empty(Optional A As vbproject) As CodeModule()
Dim C As VBComponent
Dim M As CodeModule
Dim O() As CodeModule
For Each C In PjNz(A).VBComponents
    Select Case C.Type
    Case vbext_ct_ClassModule, vbext_ct_StdModule
        Set M = C.CodeModule
        If MdIsEmpty(M) Then
            PushObj O, M
        End If
    End Select
Next
PjMdAy_Empty = O
End Function

Function PjMdAy_Std(Optional A As vbproject, Optional MdNmLik$ = "*") As CodeModule()
Dim C As VBComponent
Dim O() As CodeModule
Dim M As CodeModule
For Each C In PjNz(A).VBComponents
    If C.Type = vbext_ct_StdModule Then
        Set M = C.CodeModule
        If Not MdIsEmpty(M) Then
            PushObj O, M
        End If
    End If
Next
PjMdAy_Std = O
End Function

Function PjMdNy(Optional A As vbproject) As String()
Dim O$()
PjMdNy = ObjAyPrp(PjMdAy(A), "Name", O)
End Function

Function PjMdPtrAy(Optional A As vbproject) As LongPtr()
Dim O() As LongPtr
Dim C As VBComponent, U&
U = -1
For Each C In PjNz(A).VBComponents
    Select Case C.Type
    Case vbext_ct_StdModule, vbext_ct_ClassModule
        If C.CodeModule.CountOfLines > 0 Then
            U = U + 1
            ReDim Preserve O(U)
            O(U) = ObjPtr(C.CodeModule)
        End If
    End Select
Next
PjMdPtrAy = O
End Function

Sub PjMdPtrAy__Tst()
Dim O() As LongPtr: O = PjMdPtrAy
Stop
End Sub

Function PjNm$(Optional A As vbproject)
PjNm = PjNz(A).Name
End Function

Function PjNxtMdNm$(MdNm$, Optional Pj As vbproject)
If MdIsInPj(MdNm, Pj) Then PjNxtMdNm = MdNm: Exit Function
Dim A$(): A = PjMdNy(Pj)
PjNxtMdNm = NmNxt(MdNm, A)
End Function

Function PjNz(P As vbproject) As vbproject
If IsNothing(P) Then
    Set PjNz = PjCur
Else
    Set PjNz = P
End If
End Function

Function PjNzPth$(Pth$)
If Pth = "" Then
    PjNzPth = PjPth
Else
    PjNzPth = Pth
End If
End Function

Function PjPth$(Optional A As vbproject)
PjPth = FfnPth(PjNz(A).FileName)
End Function

Sub PjPthBrw(Optional A As vbproject)
PthBrw PjPth(A)
End Sub

Function PjRfDt(pFfn$, Optional pNmPrj$ = "", Optional pAcs As Access.Application = Nothing) As Dt
Const cSub$ = "PjRfDt"
On Error GoTo R
Dim mPrj As vbproject: If Cv_Prj(mPrj, pNmPrj, pAcs) Then ss.A 1: GoTo E
Dim iRf As VBIDE.Reference
Dim mFno As Byte: If Opn_Fil_ForOutput(mFno, pFfn, True) Then ss.A 2: GoTo E
For Each iRf In mPrj.References
    With iRf
        Write #mFno, .Name, .FullPath, .BuiltIn, .Type
    End With
Next
GoTo X
Exit Function
R: ss.R
E:
X:
    Close #mFno
End Function

Function PjRfDt__Tst()
DtBrw PjRfDt("c:\tmp\aa.txt")
Shell "notepad c:\tmp\aa.txt", vbMaximizedFocus
End Function

Function PjSrcPth$(Optional A As vbproject)
Dim Pj As vbproject: Set Pj = PjNz(A)
Dim O$
O = PjPth(Pj) & "Src\": PthEns O
O = O & Pj.Name & "\":  PthEns O
PjSrcPth = O
End Function

Function PjSrcPthBrw(Optional A As vbproject)
PthBrw PjSrcPth(A)
End Function

Sub PjSwh(Optional A As vbproject)
Dim F$: F = PjAltPjFfn(A)
Debug.Print F
Stop
Select Case FfnExt(F)
Case AppaExt, AppaExtNrm: AppaOpnPj F, Appa
'Case AppxExt: SwhFxmla F
'Case AppoExt: SwhFolk F
Case Else: Er "{ExtOf-PjFfn} is invalid", FfnExt(F)
End Select
Quit
End Sub

Function PjTyNy(Optional A As vbproject) As String()
Dim O$(), MdAy() As CodeModule
MdAy = PjMdAy(A)
If AyIsEmpty(MdAy) Then Exit Function
Dim I, Md As CodeModule
For Each I In MdAy
    Set Md = I
    PushAy O, AyAddPfx(MdTyNy(Md), MdNm(Md) & ".")
Next
PjTyNy = O
End Function
