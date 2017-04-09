Attribute VB_Name = "nDta_Ds"
Option Compare Database
Option Explicit
Type Ds
    DtAy() As Dt
End Type

Function DsAddDt(A As Ds, Dt As Dt) As Ds
If DsHasTn(A, Dt.Tn) Then Er "Ds already has {Tn}", Dt.Tn
Dim O As Ds
Dim N%: N = DsNTbl(A)
ReDim Preserve A.DtAy(N)
A.DtAy(N) = Dt
DsAddDt = A
End Function

Sub DsBrw(A As Ds)
Dim J%
For J = 0 To DsUTbl(A)
    DtBrw A.DtAy(J)
Next
End Sub

Function DsDt(A As Ds, Tn_or_Idx) As Dt
Dim O%
    Dim Tn$
    If VarIsStr(Tn_or_Idx) Then
        Tn = Tn_or_Idx
        O = DsDtIdx(A, Tn)
    Else
        O = Tn_or_Idx
    End If
DsDt = A.DtAy(O)
End Function

Function DsDtIdx%(A As Ds, Tn$)
Dim J%
For J = 0 To DsUTbl(A)
    If A.DtAy(J).Tn = Tn Then DsDtIdx = J: Exit Function
Next
DsDtIdx = -1
End Function

Function DsHasTn(A As Ds, Tn$) As Boolean
Dim J%
For J = 0 To DsUTbl(A)
    If A.DtAy(J).Tn = Tn Then DsHasTn = True: Exit Function
Next
End Function

Function DsIsEmpty(A As Ds) As Boolean
DsIsEmpty = DtSz(A.DtAy) = 0
End Function

Function DsLy(A As Ds) As String()
If DsIsEmpty(A) Then Exit Function
Dim O$()
    Dim Tny$()
    Tny = DsTny(A)
    Dim J%
    For J = 0 To UB(Tny)
        PushAy O, DtLy(DsDt(A, Tny(J)), Tny(J))
    Next
DsLy = O
End Function

Function DsNew(DtAy() As Dt) As Ds
Dim J%, O As Ds
For J = 0 To UBound(DtAy)
    O = DsAddDt(O, DtAy(J))
Next
DsNew = O
End Function

Function DsNTbl&(A As Ds)
DsNTbl = DtSz(A.DtAy)
End Function

Function DsSample1() As Ds
Dim DtAy() As Dt
DtPush DtAy, DtSample1
DtPush DtAy, DtSample2
DsSample1 = DsNew(DtAy)
End Function

Function DsTny(A As Ds) As String()
Dim J%, O$()
For J = 0 To DsUTbl(A)
    Push O, A.DtAy(J).Tn
Next
DsTny = O
End Function

Function DsUTbl%(A As Ds)
DsUTbl = DsNTbl(A) - 1
End Function

Sub DsWrt__Tst()
'1 Declare
Dim A As Ds
Dim Fcsv$

'2 Assign
Dim Tny$(): Tny = DbTny
Tny = AyExcl(Tny, "TblHasNoRec_IgnoreEr")
A = TblDs(Tny)
Fcsv = TmpFil(".csv")

'3 Calling
DsWrt A, Fcsv

FtBrw Fcsv, True
End Sub
