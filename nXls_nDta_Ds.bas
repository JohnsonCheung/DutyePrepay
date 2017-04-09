Attribute VB_Name = "nXls_nDta_Ds"
Option Compare Database
Option Explicit

Function DsCrtIdxWs(A As Ds, Wb As Workbook) As Worksheet
Dim O As Worksheet
    Set O = WbAddWsAtBeg(Wb, "Idx")
DicPutCell DsNRecDic(A), WsA1(O)
Dim J%
For J = 0 To DtUB(A.DtAy)
    
Next
Set DsCrtIdxWs = O
End Function

Sub DsCrtIdxWs__Tst()
DsCrtIdxWs DsSample1, WbNew(Vis:=True)
End Sub

Function DsNRecDic(A As Ds) As Dictionary
Dim O As New Dictionary
    Dim J%
    For J = 0 To DsUTbl(A)
        O.Add A.DtAy(J).Tn, DtNRec(A.DtAy(J))
    Next
Set DsNRecDic = O
End Function

Function DsWb(A As Ds) As Workbook
Dim O As Workbook: Set O = WbNew(Vis:=True)
    Dim Tny$(): Tny = DsTny(A)
    Dim J&
    For J = 0 To UB(Tny)
        WbAddDt O, DsDt(A, J), Tny(J)
    Next
Set DsWb = O
End Function

Sub DsWb__Tst()
DsWb DsSample1
End Sub
