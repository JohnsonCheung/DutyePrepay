Attribute VB_Name = "nDao_nObj_nRs_nDo_Rs"
Option Compare Database
Option Explicit

Function RsLin$(A As DAO.Recordset, Optional pNmFld$ = "", Optional pQ$ = "", Optional pSepChr$ = CtComma)
'Aim: FInd {oLv} from all record of first field  <pRs>.<pNmFld> of each record in {pRs} into {oLv}
Dim mAyV(): 'If RsCol(mAyV, A, pNmFld) Then ss.A 1: GoTo E
RsLin = Join_AyV(mAyV, pQ, pSepChr)
End Function

Function RsLin__Tst()
Const cSub$ = "Fnd_LvFmRs"
Dim mNmt$, mNmFld$, mLv$, mCase As Byte

For mCase = 1 To 1
    Select Case mCase
    Case 1: mNmt = "mstBrand": mNmFld = "BrandId"
    Case 2
    Case 3
    End Select
    Dim mRs As DAO.Recordset: Set mRs = CurrentDb.TableDefs(mNmt).OpenRecordset
    If Fnd_LvFmRs(mLv, mRs, mNmFld) Then Stop
    mRs.Close
    Debug.Print LpApToStr(vbLf, "mCase,mNmt,mNmFld,mLv", mCase, mNmt, mNmFld, mLv)
    Debug.Print "----"
Next
End Function

Sub RsRbr(A As DAO.Recordset, pStart As Byte, pStp As Byte)
Dim I&: I = pStart
With A
    While Not .EOF
        .Edit
        .Fields(0).Value = I: I = I + pStp
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

Sub RsRmvCummulation(pRs As DAO.Recordset, FnStrKey$, pNmFldCum$, pNmFldSet$)
Const cSub$ = "Rmv_Cummulation"
'   Output: the field pRs->pNmFldSet will be Updated
'   Input : pRs         assume it has been sorted in proper order
'           FnStrKey      is the list of key fields name used as grouping the records in pRs (records with FnStrKey value considered as a group)
'           pNmFldCum   is the value fields used to do the cummulation to set the pNmFldSet
'           pNmFldSet   is the field required to set
'   Logic:
'           For each group of records in pRs, the pNmFldSet will be set by removing cummulation in the field VFld.
'           (Note: Assuming VFld is already in cummulation)
'
If Trim(FnStrKey) = "" Then Er "FnStrKey is empty string"
Dim mAnFldKey$(): mAnFldKey = Split(FnStrKey, CtComma)
Dim NKey%: NKey = Sz(mAnFldKey)
ReDim mAyKvLas(NKey - 1)
Dim mLasRunningQty As Double
With pRs
    While Not .EOF
        If Not IsSamKey_ByAnFldKey(pRs, mAnFldKey, mAyKvLas) Then
            mLasRunningQty = 0
            Dim J%
            For J = 0 To NKey - 1
                mAyKvLas(J) = pRs.Fields(mAnFldKey(J)).Value
            Next
        End If
        .Edit
        .Fields(pNmFldSet).Value = .Fields(pNmFldCum).Value - mLasRunningQty
        mLasRunningQty = .Fields(pNmFldCum).Value
        .Update
        .MoveNext
    Wend
End With
End Sub

Function RsSel(Rs As DAO.Recordset, FnStr$) As Dictionary
'Aim: Build {oLv} by {FnStr} in {pRs}
Stop
'Dim mAnFld_Lcl$(), mAnFld_Host$(): If Brk_Lm_To2Ay(mAnFld_Lcl, mAnFld_Host, FnStr) Then ss.A 1: GoTo E
'Dim N%: N = Sz(mAnFld_Lcl)
'With Rs
'    Dim J%, mA$
'    If pIsNoNm Then
'        For J = 0 To N - 1
'            oLv = Push(oLv, Q_V(.Fields(mAnFld_Lcl(J)).Value), pSep$)
'        Next
'    Else
'        For J = 0 To N - 1
'            If Join_NmV(mA, mAnFld_Host(J), .Fields(mAnFld_Lcl(J)).Value, pBrk) Then ss.A 1: GoTo E
'            oLv = Push(oLv, mA, pSep$)
'        Next
'    End If
'End With
End Function

Sub RsSetCummulation(A As DAO.Recordset, pLoKey$, VFld$, pSetFld$)
'Aim: Set Cummulation of <VFld> into <pSetFLd> with grouping as defined in list of key fields <pKeyFlds>
'Output: the field pRs->pSetFld will be Updated
'Input : pRs, pKeyFlds, VFld, pSetFld
''pRs     : Assume it has been sorted in proper order
''pKeyFlds: a list of key fields used as grouping the records in pRs (same records with pKeyFlds value considered as a group)
''VFld : VFld is the value field name used to do the cummulation to set the pSetFld.  If ="", use 1 as value.
''pSetFld : the field required to set
'Logic : For each group of records in pRs, the pSetFld will be set to cummulate the field VFld
'Example: in ATP.mdb: ATP_35_FullSetNew_3Upd_Qty_As_Cummulate_RunCode()
''- Input table is : tmpATP_FullSetNew
''                   FGDmdId / FG / CmpSupTypSeq / CmpSupTyp / DelveryDate / Cmp / Qty / RunningQty
''- pRs      = currentdtable("tblATP_FullSetNew").openrecordset
''             pRs.index = "PrimaryKey"
''             pRs.PrimaryKey is : FGDmdId / FG / Cmp / CmpSupTypSeq / CmpSupTyp / DeliveryDate
''- pKeyFlds = FGDmdId / FG / Cmp
''- VFld  = "Qty"
''- pSetFld  = RunningQty
Dim mAnFldKey$(): mAnFldKey = Split(pLoKey, CtComma)
Dim NKey%: NKey = Sz(mAnFldKey)
ReDim mAyLasKeyVal(NKey - 1)
Dim J As Byte: For J = 0 To NKey - 1
    mAyLasKeyVal(J) = "xxxx"
Next
Dim mQ_Run As Double
With A
    While Not .EOF
        If IsSamKey_ByAnFldKey(A, mAnFldKey, mAyLasKeyVal) Then
            If VFld = "" Then
                mQ_Run = mQ_Run + 1
            Else
                mQ_Run = mQ_Run + Nz(A.Fields(VFld).Value, 0)
            End If
        Else
            If VFld = "" Then
                mQ_Run = 1
            Else
                mQ_Run = Nz(A.Fields(VFld).Value, 0)
            End If
            For J = 0 To NKey - 1
                mAyLasKeyVal(J) = A.Fields(mAnFldKey(J)).Value
            Next
        End If
        .Edit
        .Fields(pSetFld).Value = mQ_Run
        .Update
        .MoveNext
    Wend
End With
End Sub

Sub RsSetSno(A As DAO.Recordset, Optional SnoFldNm$ = "Sno")
With A
    Dim S&
    While Not .EOF
        .Edit
        S = S + 1
        .Fields(SnoFldNm).Value = S
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

Function RsToStr(oLm$, pRs As DAO.Recordset _
    , Optional pNmFld0$ = "" _
    , Optional pNmFld1$ = "" _
    , Optional pBrkChr$ = "=" _
    , Optional pSepChr$ = vbCrLf) As Boolean
'Aim: Build {oLm} from all records in {pRs} which have 2 fields {pNmFld1} & {pNmFld2}
Const cSub$ = "RsToStr"
On Error GoTo R
Dim mNmFld0$: mNmFld0 = NonBlank(pNmFld0, pRs.Fields(0).Name)
Dim mNmFld1$: mNmFld1 = NonBlank(pNmFld1, pRs.Fields(1).Name)
oLm = ""
With pRs
    While Not .EOF
        oLm = Push(oLm, .Fields(mNmFld0).Value & pBrkChr & .Fields(mNmFld1).Value, pSepChr)
        .MoveNext
    Wend
End With
Exit Function
R: ss.R
E:
End Function

Function RsToStr__Tst()
TblCrt_ByFldDclStr "#Tmp", "Itm Text 10,N Text 50,X Text 50"
If Run_Sql("Insert into [#Tmp] values ('Tbl','1,2,3',',x,xx,xxx')") Then Stop
Dim mLm$: If Set_Lm_ByTbl(mLm, "#Tmp") Then Stop
Debug.Print mLm
End Function

Function RsWrtFx(Rs As Recordset, Fx$, Optional WsNm$ = "Data")
Dim Cell As Range
    Set Cell = WsA1(WsNew(WsNm))
RsPutCell Rs, Cell
WbSavAs RgWb(Cell), Fx
End Function

Function RsWrtFx__Tst()
Dim Fx$: Fx = TmpFx
Dim Rs As Recordset: Set Rs = SqlRs("Select * from Permit")
RsWrtFx Rs, Fx
FxWb(Fx).Application.Visible = True
End Function

