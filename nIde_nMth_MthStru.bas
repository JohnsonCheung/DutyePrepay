Attribute VB_Name = "nIde_nMth_MthStru"
Option Compare Database
Option Explicit
Type MthStru
    BEIdx() As Long
    Brk As MthBrk
End Type

Function MthStruAy(Optional MthNm$, Optional PrpTy$, Optional A As CodeModule) As MthStru()
Dim ALy$()
Dim ABEIdxAy()
    Dim Md As CodeModule:
    Dim FmI&
    Set Md = MdNz(A)
    FmI = Md.CountOfDeclarationLines
    ALy = MdLy(Md)
    ABEIdxAy = MdLyToMthBEIdxAy(ALy, MthNm, PrpTy, FmI) ' By = BIdx array   ! Bidx array of given-module-A
                               
If AyIsEmpty(ABEIdxAy) Then Exit Function
Dim O() As MthStru
    Dim U%
        U = UBound(ABEIdxAy)
    ReDim O(U)
    Dim J%, M As MthStru, MthLin$, BEIdx&()
    For J = 0 To UBound(ABEIdxAy)
        BEIdx = ABEIdxAy(J)
        MthLin = SrcLyOneContinueLin(ALy, BEIdx(0))
        M.BEIdx = BEIdx
        M.Brk = MthBrkNew(MthLin)
        O(J) = M
    Next
MthStruAy = O
End Function

Function MthStruAyIsEmpty(A() As MthStru) As Boolean
On Error GoTo Y
Dim U&: U = UBound(A)
Exit Function
Y: MthStruAyIsEmpty = True
End Function

Function MthStruAyKeyAy(A() As MthStru) As String()
If MthStruAyIsEmpty(A) Then Exit Function
Dim U&: U = UBound(A)
Dim O$(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = MthStruKey(A(J))
Next
MthStruAyKeyAy = O
End Function

Sub MthStruBrw(A As MthStru)
DtBrw MthStruDt(A)
End Sub

Function MthStruDt(A As MthStru) As Dt
With A.Brk
    Dim D()
    D = DrAyNewKVByFldAp("BIdx EIdx MthNm Mdy MthTy PrpTy RetTyChr RetAs PrmStr", _
            A.BEIdx(0), A.BEIdx(1), .Nm, .mFY, .Ty, .PrpTy, .RetTyChr, .RetAs, .PrmStr)
End With
MthStruDt = DtNew(LvsSplit("MthStruPrp Val"), D)
End Function

Function MthStruIsEmpty(A As MthStru) As Boolean
MthStruIsEmpty = A.Brk.Nm = ""
End Function

Function MthStruKey$(A As MthStru)
Dim M$, N$, P$
With A.Brk
    P = .PrpTy
    N = .Nm
    Select Case .mFY
    Case "", "Public": M = 0
    Case "Friend": M = 1
    Case "Private": M = 2
    Case Else: Er "{Mfy} is invalid", .mFY
    End Select
End With
MthStruKey = FmtQQ("?:?:?", M, N, P)
End Function

