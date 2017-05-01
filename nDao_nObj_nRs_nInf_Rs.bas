Attribute VB_Name = "nDao_nObj_nRs_nInf_Rs"
Option Compare Database
Option Explicit

Sub RsCls(A As DAO.Recordset)
On Error Resume Next
A.Close
End Sub

Function RsCmpFrm(oIsSam As Boolean, pRs As DAO.Recordset, pFrm As Access.Form, FnStr$) As Boolean
'Aim: Compare {pRs} Field with OldValue of control in {pFrm} by using list of name in {FnStr}
'     {FnStr} has aaa=xxx,bbb,ccc format, aaa,bbb,ccc are the Form's name and xxx,bbb,ccc are the Rs field name.
Const cSub$ = "RsCmpFrm"
oIsSam = False
On Error GoTo R
Dim mAn_Frm$(), mAn_Rs$(): If Brk_Lm_To2Ay(mAn_Frm, mAn_Rs, FnStr) Then ss.A 1: GoTo E
Dim mA$, mNmFld_Frm$, mNmFld_Rs$, mIsEq As Boolean
Dim J%, N%
N = Sz(mAn_Frm)
For J = 0 To N - 1
    mNmFld_Frm = mAn_Frm(J)
    mNmFld_Rs = mAn_Rs(J)
    Dim mV_FrmOld: mV_FrmOld = pFrm.Controls(mNmFld_Frm).OldValue
    Dim mV_Rs: mV_Rs = pRs.Fields(mNmFld_Rs).Value
    If IfEq(mIsEq, mV_FrmOld, mV_Rs) Then ss.A 1: GoTo E
    If Not mIsEq Then mA = Push(mA, Fmt("Rs({0})=[{1}] Frm({2}).New=[{3}] .Old=[{4}]", mNmFld_Rs, mV_Rs, mNmFld_Frm, pFrm.Controls(mNmFld_Frm).Value, mV_FrmOld), vbCrLf)
Next
If mA <> "" Then ss.A 1, "There is some fields OldValue not same as the host", "The fields", mA: Exit Function
oIsSam = True
Exit Function
R: ss.R
E:
End Function

Function RsCol(A As DAO.Recordset, Optional FldNm$) As Variant
'Aim: Find the first field in {pRs} for each record in pRs into {oAyV}
Dim I&
    If FldNm = "" Then
        I = 1
    Else
        I = A.Fields(FldNm).OrdinalPosition
    End If
Dim O()
With A
    Dim N&: N = 0
    While Not .EOF
        ReDim Preserve O(N)
        O(N) = .Fields(I).Value
        N = N + 1
        .MoveNext
    Wend
End With
RsCol = O
End Function

Function RsCol_Int(A As DAO.Recordset, Optional FldNm$) As Integer()
RsCol_Int = AyIntAy(RsCol(A, FldNm))
End Function

Function RsDic(A As DAO.Recordset) As Dictionary
Dim J&
Dim O As New Dictionary
Dim I As DAO.Field
For Each I In A.Fields
    O.Add I.Name, I.Value
Next
Set RsDic = O
End Function

Function RsDr(Rs As Recordset, Optional FstNFld%) As Variant()
RsDr = FldsDr(Rs.Fields, FstNFld)
End Function

Function RsDrAy(Rs As Recordset, Optional FstNFld%) As Variant()
Dim O()
    With Rs
        While Not .EOF
            Push O, RsDr(Rs, FstNFld)
            .MoveNext
        Wend
    End With
RsDrAy = O
End Function

Function RsDt(A As Recordset, Optional FstNFld%, Optional Tn = "Rs") As Dt
Dim D() As Variant
Dim F$()
F = RsFny(A, FstNFld)
D = RsDrAy(A, FstNFld)
RsDt = DtNew(F, D, CStr(Tn))
End Function

Function RsFny(A As Recordset, Optional FstNFld%) As String()
RsFny = FldsFny(A.Fields, FstNFld)
End Function

Function RsIdx(A As Recordset, Fny$()) As Long()
RsIdx = AySubsetIdxAy(RsFny(A), Fny)
End Function

Function RsIsEmptyRec(A As Recordset) As Boolean
RsIsEmptyRec = DrIsEmptyRec(RsDr(A))
End Function

Function RsIsEq(Rs1 As DAO.Recordset, Rs2 As DAO.Recordset) As Boolean
If Rs1.Fields.Count <> Rs2.Fields.Count Then Exit Function
RsIsEq = DicIsEq(RsDic(Rs1), RsDic(Rs2))
End Function

Function RsIsEq__Tst()
Dim Rs1 As Recordset, Rs2 As Recordset
Set Rs1 = CurrentDb.OpenRecordset("Select * from Permit")
Set Rs2 = CurrentDb.OpenRecordset("Select * from Permit")
Debug.Assert RsIsEq(Rs1, Rs2) = True
End Function

Function RsIsMatch(A As DAO.Recordset, Dic As Dictionary) As Boolean
'Aim: Return if the list of fields of name in {pKeyFlds} of current record of {A} has the same values as in {pLastKey}
Dim U%: U = Dic.Count - 1
Dim J%, I
For Each I In Dic
    If CStr(A.Fields(I).Value) <> CStr(Dic(I)) Then
        Exit Function
    End If
Next
RsIsMatch = True
End Function

Function RsIsMatch__Tst()
Dim Rs As Recordset
Dim Dic As Dictionary
    Set Rs = TblRs("Permit")
    Set Dic = DicNew("Permit=3")

With Rs
    While Not .EOF
        If RsIsMatch(Rs, Dic) Then Stop
        .MoveNext
    Wend
    .Close
End With

End Function

Function RsIsMatchKey(A As DAO.Recordset, KeyVal()) As Boolean
'Aim: Return True if first N fields of {Rs} is same as KeyVal
Dim U%: U = UB(KeyVal)
Dim J%
For J = 0 To U
    If A.Fields(J).Value <> KeyVal(J) Then Exit Function
Next
RsIsMatchKey = True
End Function

Function RsIsSubSet(oIsSubSet As Boolean, pRsSub As DAO.Recordset, pRsSuper As DAO.Recordset) As Boolean
'Aim: Compare each field in {pRsSub} is in {pRsSub} and have same value
Const cSub$ = "RsIsSubSet"
On Error GoTo R
Dim mIsEq As Boolean
Dim J%: For J = 0 To pRsSub.Fields.Count - 1
    With pRsSub.Fields(J)
        Dim mNm$: mNm = .Name
        Dim mV1: mV1 = .Value
        Dim mTyp As DAO.DataTypeEnum: mTyp = .Type
    End With
    If IfEq(mIsEq, mV1, pRsSuper.Fields(mNm).Value) Then ss.A 1, , , "Field with IsEq err", mNm: GoTo E
    If Not mIsEq Then GoTo E
Next
Exit Function
R: ss.R
E:
End Function

Function RsSq(A As Recordset)
RsSq = DtSq(RsDt(A))
End Function

Function RsWsNew(A As DAO.Recordset, Optional WsNm$, Optional NoLo As Boolean) As Worksheet
Dim O As Worksheet
Set O = WsNew(WsNm)
DtPutCell RsDt(A), WsA1(O), NoLo
Set RsWsNew = O
End Function

Function RsWsNew__Tst()
Dim Rs As DAO.Recordset
Set Rs = CurrentDb.OpenRecordset("Select * from Permit")
WsVis RsWsNew(Rs)
End Function
