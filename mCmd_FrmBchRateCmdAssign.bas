Attribute VB_Name = "mCmd_FrmBchRateCmdAssign"
Option Compare Database
Option Explicit

Sub FrmBchRateCmdAssign(xYY As Byte, xMM As Byte, xDD As Byte)
'Aim: For each OH of 'new' batch#, try assign to PermitD->BchNo & create record in SkuB
'     - For those OH batch# cannot find a record in PermitD->(Sku+BchNo), try assign these BchNo to PermitD->BchNo.
'     - When assigning batch#
'       - Sum of PermitD->Qty for Sku+BchNo > OH
'       - All PermitD->Rate of same Sku+BchNo are the same.
'       - For PermitD, use those latest PermitDate first.
'       - For OH     , use those smallest Batch# first.
'Ref: SkuB = Sku BchNo | DutyRateB
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
'     Permit  = * *No | *Date PostDate NSku Qty Tot GLAc GLAcName BankCode ByUsr DteCrt DteUpd
'     OH      = YY MM DD YpStk Sku BchNo | Bott Val
'SqlRun "Update PermitD set BchNo=Null"
RR_1CrtTmpCurrentDb
With CurrentDb.OpenRecordset("Select * from `#Assign_OH` order by Sku,BchNo")      ' = Sku,BchNo,OH
    While Not .EOF
        RR_2BchCurrentDb .Fields(0).Value, .Fields(1).Value, .Fields(2).Value
        .MoveNext
    Wend
    .Close
End With
TblSkuBRfh  ' Re-Create record in SkuB
zBldRecordSourceTable xYY, xMM, xDD
End Sub

Sub FrmBchRateCmdAssign__Tst()
FrmBchRateCmdAssign 17, 3, 31
End Sub

Sub xxCurrentDb()
'Aim: Update PermitD->BchNo & create record in SkuB
'     - Each Tax-Paid-OH item's batch# will assign to PermitD->BchNo
'     - Sum of PermitD->Qty for Sku+BchNo > OH
'     - All PermitD->Rate of same Sku+BchNo are the same.
'     - Use those with PermitDate is latest first.
'Ref: SkuB = Sku BchNo | DutyRateB
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
'     Permit  = * *No | *Date PostDate NSku Qty Tot GLAc GLAcName BankCode ByUsr DteCrt DteUpd
'     OH      = YY MM DD YpStk Sku BchNo | Bott Val
DoCmd.SetWarnings False
xx_1CrtTmpCurrentDb
With CurrentDb.OpenRecordset("#Assign_SKU")
    While Not .EOF
        SysCmd acSysCmdSetStatus, .Fields(0).Value
        xx_2SKUCurrentDb .Fields(0).Value
        .MoveNext
    Wend
    SysCmd acSysCmdClearStatus
    .Close
End With
End Sub

Sub zBldRecordSourceTable(xYY As Byte, xMM As Byte, xDD As Byte)
'Aim: Refresh table-(frmBchRate frmBchRateOH) from OH (Co=8600)
'Refresh frmBchRate = Sku DesSku OH IsNoAssign DutyRate BottPerAc DutyRateBott
'Refresh frmBchRateOH = Sku BchNo OH IsNoAssign

'Refresh table-frmBchRateOH ------------------------------------------
    SqlRun "Delete from frmBchRateOH"
    
    SqlRunQQ "Insert into frmBchRateOH (Sku,BchNo,OH,IsNoAssign)" & _
    " Select Distinct SKU, BchNo, Sum(Bott), False" & _
    " from OH x inner join YpStk a on x.YpStk=a.YpStk" & _
    " where IsTaxPaid and YY=? and MM=? and DD=? and Co=1" & _
    " and SKU in (Select SKU from SKU_StkHld where TaxRate is not null)" & _
    " group by SKU,BchNo" & _
    " having Sum(x.bott)<>0", _
    xYY, xMM, xDD    ' 'Co=1 is 8600 HK
    
    SqlRun "Select Sku,BchNo,Sum(Qty) as QtyPermit into `#frmBchRate_Permit` from PermitD where Nz(BchNo,'')<>'' Group by Sku,BchNo"
    SqlRun "Update frmBchRateOH x left join `#frmBchRate_Permit` a on a.Sku=x.Sku and a.BchNo=x.BchNo set IsNoAssign=True where OH>Nz(QtyPermit,0)"
    TblDrp "#frmBchRate_Permit"

'Refresh table-frmBchRate ------------------------------------------
    SqlRun "Delete from frmBchRate"
    SqlRun "Insert into frmBchRate (Sku,OH,IsNoAssign) Select Distinct Sku, Sum(OH), Min(IsNoAssign) from frmBchRateOH group by Sku"
    SqlRun "Update frmBchRate x inner join Sku_StkHld a on a.Sku=x.Sku set x.DutyRate=a.TaxRate"
    SqlRun "Update frmBchRate x inner join tblSKU a on a.SkuTxt=x.Sku set x.BottPerAc=a.BtlPerCs"
    SqlRun "Update (frmBchRate x inner join Sku88 a on a.Sku=x.Sku) inner join tblSku b on b.SkuTxt=a.Sku88 set x.BottPerAc=b.BtlPerCs"
    SqlRun "Update frmBchRate set DutyRateBott=DutyRate/BottPerAc where Nz(BottPerAc,0)<>0"
    SqlRun "Update frmBchRate x inner join qSku a on a.Sku=x.Sku set DesSku=`Sku Description`"
End Sub

Private Sub RR_1CrtTmpCurrentDb()
'Aim: Create #Assign_OH from table-OH for those Tax-Paid item with SKU+BchNo not found in table PermitD
'Ref  #Assign_OH  = SKU,BchNo,OH
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
Dim D As YMD: D = TblOHMaxYMD

TblDrp "#Assign_OH"
SqlRun Fmt_Str("Select Distinct SKU,BchNo,Sum(Bott) as OH" & _
" into `#Assign_OH` from OH x" & _
" where YY={0} and MM={1} and DD={2}" & _
" and YpStk in (Select YpStk from YpStk where IsTaxPaid)" & _
" and SKU in (Select SKU from SKU_StkHld where TaxRate is not null)" & _
" group by SKU,BchNo" & _
" having Sum(Bott)<>0" & _
" order by BchNo desc", D.Y, D.M, D.D)
SqlRun "Update`#Assign_OH` x inner join PermitD a on a.Sku=x.Sku and a.BchNo=x.BchNo set x.OH=Null"
SqlRun "Delete from `#Assign_OH` where OH is null"
DoCmd.OpenTable "#Assign_OH"
End Sub

Private Sub RR_2BchCurrentDb(pSku$, pBchNo$, pOH&)
'Aim: Read PermitD/Permit
'     Find one or more record in PermitD covering the quantity
'     Set PermitD->BchNo & Write to SkuB
'Ref:#Assign_OH  = SKU,BchNo,OH
'    PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd

'                                                                    Read PermitD in PermitDate Desc for those BchNo=''
Dim aPermitD&(), aRate@(), aQty&(): RR_2BchCurrentDb_1PermitD pSku, aPermitD, aRate, aQty: If Sz(aPermitD) = 0 Then Exit Sub
Dim J%
Dim mAyIdx%()
Dim mAyPermitD$()
Dim MRate@
mAyPermitD = RR_2BchCurrentDb_3AyPermitD(pOH, aPermitD, aRate, aQty, MRate) ' Find continuous PermitD's of same rate with quantity can cover bOH(J).
'                                                                  ' After found, return mAyPermitD, Rate, oIdx and set aQty to zero
If Sz(mAyPermitD) > 0 Then
    SqlRun Fmt_Str("Update PermitD set BchNo='{0}' where PermitD in ({1}) and SKU='{2}'", pBchNo, Join(mAyPermitD, ","), pSku)
    SqlRun Fmt_Str("Insert into SkuB (Sku,BchNo,DutyRateB) values ('{0}','{1}',{2})", pSku, pBchNo, MRate)
End If
End Sub

Private Sub RR_2BchCurrentDb_1PermitD(pSku$, oPermitD&(), ORate@(), oQty&())
'Aim: Obtain the o* from #Assign_PermitD in MinPermitDate Desc
Dim mN%
With CurrentDb.OpenRecordset(Fmt_Str("Select PermitD,Round(x.Rate,2) as Rate,x.Qty from PermitD x inner join Permit a on a.Permit=x.Permit" & _
                                     " where SKU='{0}' and Nz(BchNo,'')='' order by PermitDate Desc, x.Rate Desc, x.Qty Desc", pSku))
    While Not .EOF
        ReDim Preserve oPermitD(mN)
        ReDim Preserve ORate(mN)
        ReDim Preserve oQty(mN)
        oPermitD(mN) = !PermitD
        ORate(mN) = !Rate
        oQty(mN) = !Qty
        mN = mN + 1
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Function RR_2BchCurrentDb_3AyPermitD(pOH&, pPermitD&(), pRate@(), pQty&(), ByRef ORate@) As String()
'Aim: Find continuous PermitD's of same rate with quantity can cover bOH(J).
'     After found, return mAyPermitD, oRate, oIdx
Dim mAyIdx%(): mAyIdx = RR_2BchCurrentDb_3AyPermitD_1AyIdx(pOH, pRate, pQty)
Dim mUB%: mUB = Sz(mAyIdx) - 1: If mUB < 0 Then Exit Function
ORate = pRate(mAyIdx(0))
Dim O$()
ReDim O(mUB)
Dim I%
For I = 0 To UBound(mAyIdx)
    O(I) = pPermitD(mAyIdx(I))
Next
RR_2BchCurrentDb_3AyPermitD = O
End Function

Private Function RR_2BchCurrentDb_3AyPermitD_1AyIdx(pOH&, pRate@(), pQty&()) As Integer()
'Aim: Find continuous records of same rate with quantity can cover bOH(J).
'     After found, return oAyIdx and set aQty to zero

Dim mIdx%: mIdx = RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx(pOH, pRate, pQty) ' Find oIdx from which onward, the pQty can cover pOH and having same rate.
If mIdx = -1 Then Exit Function
Dim MRate@: MRate = pRate(mIdx)
Dim O%()
Dim J%
Dim mOH&: mOH = pOH
For J = mIdx To UBound(pQty)
    If pRate(J) <> MRate Then RR_2BchCurrentDb_3AyPermitD_1AyIdx = O: Exit Function
    mOH = mOH - pQty(J)
    If pQty(J) > 0 Then Push O, J
    If mOH <= 0 Then RR_2BchCurrentDb_3AyPermitD_1AyIdx = O: Exit Function
Next
Stop ' impossible to reach here.
End Function

Private Function RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx%(pOH&, pRate@(), pQty&())
'Aim: ' Find oIdx from which onward, the oQty can cover pOH and having same rate.
Dim O%: O = RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt(pQty, pRate)
While O <> -1
    Dim MRate@: MRate = pRate(O)
     If RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_2IsOK(pOH, O, pRate, pQty) Then RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx = O: Exit Function
    O = RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt(pQty, pRate, O + 1, MRate)
Wend
RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx = -1
End Function

Private Function RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt%(pQty&(), pRate@(), Optional pBeg% = 0, Optional p_Rate@ = 0)
'Aim: Start from pBeg, find oIdx so that pQty(oIdx)>0 and pRate(oIdx)<>p_Rate
Dim O%
For O = pBeg To UBound(pQty)
    If pQty(O) <> 0 Then
        If p_Rate <> pRate(O) Then RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt = O: Exit Function
    End If
Next
RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt = -1
End Function

Private Function RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_2IsOK(pOH&, pIdx%, pRate@(), pQty&()) As Boolean
'Return true if pIdx is coRRCurrentDbect index:
'Use pIdx & pRate() to find mRate
'From pIdx onward, pQty of records of same Rate can cover pOH return true else false
Dim MRate@: MRate = pRate(pIdx)

Dim J%
Dim mOH&: mOH = pOH
For J = pIdx To UBound(pRate)
    If pRate(J) <> MRate Then Exit Function
    If pQty(J) >= mOH Then RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_2IsOK = True: Exit Function
    mOH = mOH - pQty(J)
Next
End Function

Private Sub xx_1CrtTmpCurrentDb()
'Aim: Create 4 temp table
'     #Assign_OH` latest Tax-Paid-Item OH from table-OH
'     #Assign_SKU unique SKU from #OH
'     #Assign_Lot  Each from PermitD having same Rate and in seq of date
'     #Assign_LotD Each lot is what PermitD
'Ref  #Assign_OH  = SKU,BchNo,OH
'     #Assign_SKU = SKU
'     #Assign_Lot = Lot SKU Rate MinPermitDate | Qty
'     #Assign_LotD= Lot PermitD |
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
'     Permit  = * *No | *Date ...
Dim D As YMD: D = TblOHMaxYMD

SqlRun Fmt_Str("Select Distinct SKU,BchNo,Sum(Bott) as OH into `#Assign_OH` from OH x" & _
" where YY={0} and MM={1} and DD={2}" & _
" and YpStk in (Select YpStk from YpStk where IsTaxPaid)" & _
" and SKU in (Select SKU from SKU_StkHld where IfTaxable='Y')" & _
" group by SKU,BchNo order by BchNo desc ", D.Y, D.M, D.D)
SqlRun Fmt_Str("Select Distinct SKU into `#Assign_SKU` from `#Assign_OH`")

xDlt.Dlt_Tbl "#Assign_Lot"
xDlt.Dlt_Tbl "#Assign_LotD"
CurrentDb.Execute "Create Table `#Assign_Lot` (Lot Integer, Sku Text(15), Rate Currency, MinPermitDate date, Qty Long," & _
" Constraint PrimaryKey Primary Key (Lot), Constraint `#Assign_Lot` Unique (Sku,Rate,MinPermitDate)) "
CurrentDb.Execute "Create Table `#Assign_LotD` (Lot Integer, PermitD Long, Constraint `#Assign_LotD` unique (Lot,PermitD))"

Dim mLot%: mLot = 1
Dim mLasSku$
Dim mLasRate$
Dim mMinPermitDate As Date
Dim mQty&
Dim mRsLot As Recordset:  Set mRsLot = CurrentDb.TableDefs("#Assign_Lot").OpenRecordset
Dim mRsLotD As Recordset: Set mRsLotD = CurrentDb.TableDefs("#Assign_LotD").OpenRecordset
Dim mRs As Recordset:     Set mRs = CurrentDb.OpenRecordset("Select x.PermitD, Sku, Round(x.Rate,2) as Rate, PermitDate, x.Qty from PermitD x inner join Permit a on a.Permit=x.Permit order by Sku,PermitDate,Rate")
With mRs
    mRsLotD.AddNew
    mRsLotD!Lot = mLot
    mRsLotD!PermitD = !PermitD
    mRsLotD.Update
    
    mLasSku = !Sku
    mLasRate = !Rate
    mQty = !Qty
    mMinPermitDate = !PermitDate
    .MoveNext
    While Not .EOF
        If mLasSku <> !Sku Or mLasRate <> !Rate Then
            mRsLot.AddNew
            mRsLot!Lot = mLot
            mRsLot!Sku = mLasSku
            mRsLot!Rate = mLasRate
            mRsLot!MinPermitDate = mMinPermitDate
            mRsLot!Qty = mQty
            mRsLot.Update
            
            mLot = mLot + 1
            
            mLasSku = !Sku
            mLasRate = !Rate
            mQty = !Qty
            mMinPermitDate = !PermitDate
        Else
            mQty = mQty + !Qty
        End If
        
        mRsLotD.AddNew
        mRsLotD!Lot = mLot
        mRsLotD!PermitD = !PermitD
        mRsLotD.Update
        
        .MoveNext
    Wend
    mRsLot.AddNew
    mRsLot!Lot = mLot
    mRsLot!Sku = mLasSku
    mRsLot!Rate = mLasRate
    mRsLot!MinPermitDate = mMinPermitDate
    mRsLot!Qty = mQty
    mRsLot.Update
End With
End Sub

Private Sub xx_2SKUCurrentDb(pSku$)
If pSku = "1034125" Then Stop
'Aim: Read *OH & *Lot
'     For each *OH find one or more *Lot covering the quantity
'     Set PermitD->BchNo & Write to SkuB
'Ref:#Assign_OH  = SKU,BchNo,OH
'    #Assign_Lot = Lot SKU Rate MinPermitDate | Qty
'    #Assign_LotD= Lot PermitD |
'    PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd

Dim bBchNo$(), bOH&():          xx_2SKUCurrentDb_1OH pSku, bBchNo, bOH
Dim aLot%(), aRate@(), aQty&(): xx_2SKUCurrentDb_2Lot pSku, aLot, aRate, aQty: If Sz(aLot) = 0 Then Exit Sub
Dim cPermitD$
Dim J%
Dim mLotIdx%
For J = 0 To UBound(bOH)
    mLotIdx = xx_2SKUCurrentDb_3LotIdx(bOH(J), aLot, aQty) ' Find those Lot's quantity can cover bOH(J)
    If mLotIdx >= 0 Then
        Dim mPermitD$(): mPermitD = SqlSy("Select PermitD from `#Assign_LotD` where Lot=" & aLot(mLotIdx))
        SqlRun Fmt_Str("Update PermitD set BchNo='{0}' where PermitD in ({1}) and SKU='{2}'", bBchNo(J), Join(mPermitD, ","), pSku)
        SqlRun Fmt_Str("Insert into SkuB (Sku,BchNo,DutyRateB) values ('{0}','{1}',{2})", pSku, bBchNo(J), aRate(mLotIdx))
'    Else
'        Stop
    End If
Next
End Sub

Private Sub xx_2SKUCurrentDb_1OH(pSku$, oBchNo$(), oOH&())
Dim mN%
mN = 0
With CurrentDb.OpenRecordset(Fmt_Str("Select SKU,BchNo,OH from `#Assign_OH` where SKU='{0}' order by OH Desc", pSku))
    While Not .EOF
        ReDim Preserve oBchNo(mN)
        ReDim Preserve oOH(mN)
        oBchNo(mN) = !BchNo
        oOH(mN) = !OH
        mN = mN + 1
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub xx_2SKUCurrentDb_2Lot(pSku$, oLot%(), ORate@(), oQty&())
'Aim: Obtain the o* from #Assign_Lot in MinPermitDate Desc
'Ref: #Assign_Lot = Lot SKU Rate MinPermitDate | Qty
Dim mN%
With CurrentDb.OpenRecordset(Fmt_Str("Select Lot,Rate,Qty from `#Assign_Lot` where SKU='{0}' order by MinPermitDate Desc, Qty Desc", pSku))
    While Not .EOF
        ReDim Preserve oLot(mN)
        ReDim Preserve ORate(mN)
        ReDim Preserve oQty(mN)
        oLot(mN) = !Lot
        ORate(mN) = !Rate
        oQty(mN) = !Qty
        mN = mN + 1
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Function xx_2SKUCurrentDb_3LotIdx%(pOH&, pLot%(), pQty&())
'Aim: Return IdxOf pLot()/pQty() which pQty can cover pOH
Dim O%()
Dim J%
For J = 0 To UBound(pQty)
    If pQty(J) >= pOH Then xx_2SKUCurrentDb_3LotIdx = J: Exit Function
Next
xx_2SKUCurrentDb_3LotIdx = -1
End Function

