Attribute VB_Name = "nDao_nTbl_nInf_Tbl"
Option Compare Database
Option Explicit

Function TblCnnStr$(T$, Optional A As database)
TblCnnStr = DbNz(A).TableDefs(T).Connect
End Function

Sub TblDt__Tst()
DtBrw TblDt("Permit")
End Sub

Function TblFldAy(T, Optional Fny, Optional Db As database) As DAO.Field()
Dim D As database: Set D = DbNz(Db)
If IsMissing(Fny) Then
    TblFldAy = FldsFldAy(D.TableDefs(T).Fields)
Else
    Dim U&: U = UB(Fny)
    Dim O() As DAO.Field: ReSz O, U
    With D.TableDefs(T)
        Dim J%
        For J = 0 To U
            Set O(J) = .Fields(Fny(J))
        Next
    End With
End If
TblFldAy = O
End Function

Function TblFldToLng&(T, F, Where$, Optional A As database)
TblFldToLng = TblFldV(T, F, Where, A)
End Function

Function TblFldV(T, F, Where$, Optional A As database)
TblFldV = SqlV(SqsOfSel(T, F, Where))
End Function

Function TblFny(T, Optional FstNFld%, Optional A As database) As String()
TblFny = FldsFny(DbNz(A).TableDefs(T).Fields, FstNFld)
End Function

Function TblHasFld(T, F, Optional A As database) As Boolean
Dim Flds As DAO.Fields: Set Flds = DbNz(A).TableDefs(T).Fields
TblHasFld = FldsHasFld(Flds, F)
End Function

Function TblHasIdx(T, IdxNm$, Optional A As database, Optional pSilient As Boolean = False) As Boolean
On Error GoTo R
Dim Nm$: Nm = DbNz(A).TableDefs(T).Indexes(IdxNm).Name
TblHasIdx = True
Exit Function
R:
End Function

Function TblHasNoRec(T$, Optional A As database) As Boolean
TblHasNoRec = TblNRec(T, A) = 0
End Function

Function TblHasNoRec_IgnoreEr(T$, Optional A As database) As Boolean
TblHasNoRec_IgnoreEr = TblNRec_IgnoreEr(T, A) = 0
End Function

Function TblIsLnk(T$, Optional A As database) As Boolean
On Error GoTo R
TblIsLnk = DbNz(A).TableDefs(T).Connect <> ""
Exit Function
R:
End Function

Function TblNonPKAy(T$, Optional A As database) As String()
Dim A1$(): A1 = TblFny(T, , A)
Dim A2$(): A2 = TblPKAy(T, A)
TblNonPKAy = AyMinus(A1, A2)
End Function

Function TblNonPKStr$(T$, D As database)
TblNonPKStr = FnyToStr(TblNonPKAy(T, D))
End Function

Function TblNRec&(T$, Optional A As database)
TblNRec = SqlLng(SqsOfSel(T, "Count(*)"), A)
End Function

Function TblNRec_IgnoreEr&(T$, Optional A As database)
On Error Resume Next
TblNRec_IgnoreEr = TblNRec(T, A)
End Function

Function TblPIdx(T$, D As database) As DAO.Index
On Error Resume Next
Dim I As Index
For Each I In D.TableDefs(T).Indexes
    If I.Primary Then Set TblPIdx = I: Exit Function
Next
End Function

Function TblPKAy(T$, D As database) As String()
Dim I As DAO.Index: Set I = TblPIdx(T, D): If IsNothing(I) Then Exit Function
Dim O$()
Dim F
For Each F In TblPIdx(T, D).Fields
    Push O, F.Name
Next
TblPKAy = O
End Function

Function TblPKStr$(T$, D As database)
TblPKStr = FnyToStr(TblPKAy(T, D))
End Function

Function TblPriIdx(T$, Optional A As database) As Index
On Error GoTo X
Dim I As Index
For Each I In DbNz(A).TableDefs(T).Indexes
    If I.Primary Then Set TblPriIdx = I: Exit Function
Next
X:
End Function

Function TblRs(T, Optional A As database) As Recordset
Set TblRs = SqlRs((SqsOfSel(T)), A)
End Function

Function TblStru$(T$, Optional D As database)
Dim Db As database: Set Db = DbNz(D)
Dim P$: P = TblPKStr(T, Db)
Dim R$: R = TblNonPKStr(T, Db)
Dim A$: A = FmtQQ(" = ? | ?", P, R)
TblStru = T & Replace(A, T, "*")
End Function

Function TblWs(T, Optional WsNm$, Optional A As database) As Worksheet
Dim O As Worksheet
Set O = WsNew(WsNm)
DtPutCell TblDt(T, , A), WsA1(O)
Dim WsNm1$: WsNm1 = IIf(WsNm = "", T, WsNm)
O.Name = WsNm1
Set TblWs = O
End Function
