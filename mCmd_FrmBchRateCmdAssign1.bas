Attribute VB_Name = "mCmd_FrmBchRateCmdAssign1"
Option Compare Database
Option Explicit
Const C_TmpAssignOH = "#Assign_OH"
Const C_TmpAssignSKU = "#Assign_SKU"
Const C_TmpAssignLot = "#Assign_Lot"
Const C_TmpAssignLotD = "#Assign_LotD"
Const C_InpYpStk = "YpStk"
Const C_InsSkuB = "SkuB"

Sub FrmBchRateCmdAssignBchNo(xYY As Byte, xMM As Byte, xDD As Byte)
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
Tmp1AssignOH_Bld
Dim Sku$, BchNo$, OH&, Dr
Dim S$: S = FmtQQ("Select SKU,BchNo,OH from `?` order by Sku,BchNo", C_TmpAssignOH)
For Each Dr In SqlDrAy(S)
    AyAsg Dr, Sku, BchNo, OH
    RR_2BchCurrentDb Sku, BchNo, OH
Next
TblSkuBRfh ' Re-Create record in SkuB
zBldRecordSourceTable xYY, xMM, xDD
End Sub

Sub FrmBchRateCmdAssignBchNo__Tst()
FrmBchRateCmdAssignBchNo 17, 3, 31
End Sub

Private Sub AInpOH()
ZOpn "OH"
End Sub

Private Sub AInpOHMaxYMD()
MsgBox "TblOHMaxYMD" & vbCrLf & YMDToStr(TblOHMaxYMD)
End Sub

Private Sub AInpSKU_StkHld()
ZOpn "SKU_StkHld"
End Sub

Private Sub AInpYpStk()
ZOpn "YpStk"
End Sub

Private Sub AOupPermitD()
ZOpn "PermitD"
End Sub

Private Sub AOupSkuB()
ZOpn "SkuB"
End Sub

Private Sub ATmpAssignLot()
ZOpn C_TmpAssignLot
End Sub

Private Sub ATmpAssignLotD()
ZOpn C_TmpAssignLotD
End Sub

Private Sub ATmpAssignOH()
ZOpn "#Assign_OH"
End Sub

Private Sub ATmpAssignSKU()
ZOpn C_TmpAssignSKU
End Sub

Private Sub OupUpdPermitD_and_InsSkuB()
'Aim: Update PermitD->BchNo & create record in SkuB
'     - Each Tax-Paid-OH item's batch# will assign to PermitD->BchNo
'     - Sum of PermitD->Qty for Sku+BchNo > OH
'     - All PermitD->Rate of same Sku+BchNo are the same.
'     - Use those with PermitDate is latest first.
'Ref: SkuB = Sku BchNo | DutyRateB
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
'     Permit  = * *No | *Date PostDate NSku Qty Tot GLAc GLAcName BankCode ByUsr DteCrt DteUpd
'     OH      = YY MM DD YpStk Sku BchNo | Bott Val

xx_1CrtTmpCurrentDb
Dim SkuAy$()
    SkuAy = SqlSy(FmtQQ("Select Sku from [?]", C_TmpAssignSKU))
Dim OUpdPermitD$()
Dim OInsSkuB$()
    If AyIsEmpty(SkuAy) Then Exit Sub
    Dim SqlU$(), SqlI$(), Sku, Dr
    For Each Sku In SkuAy
        SysCmd acSysCmdSetStatus, Sku
        Dr = Upd_and_Ins_SqlAy(Sku)
        AyAsg Dr, SqlU, SqlI
        PushAy OUpdPermitD, SqlU
        PushAy OInsSkuB, SqlI
        SysCmd acSysCmdClearStatus
    Next
ZRunAy OUpdPermitD
ZRunAy OInsSkuB
End Sub

Private Sub RR_2BchCurrentDb(Sku$, BchNo$, OH&)
'Aim: Read PermitD/Permit
'     Find one or more record in PermitD covering the quantity
'     Set PermitD->BchNo & Write to SkuB
'Ref:#Assign_OH  = SKU,BchNo,OH
'    PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd

'                                                                    Read PermitD in PermitDate Desc for those BchNo=''
Dim PermitDAy&(), RateAy@(), QtyAy&():
    Dim DrAy()
    DrAy = SqlDrAy(FmtQQ("Select PermitD, Round(x.Rate,2) as Rate, x.Qty from PermitD x inner join Permit a on a.Permit=x.Permit" & _
                                     " where SKU='?' and Nz(BchNo,'')='' order by a.PermitDate Desc, x.Rate Desc, x.Qty Desc", Sku))
                                     
    If Sz(DrAy) = 0 Then Exit Sub
    DrAyAsg DrAy, PermitDAy, RateAy, QtyAy

Dim Rate@
PermitDAy = RR_2BchCurrentDb_3AyPermitD(OH, PermitDAy, RateAy, QtyAy, Rate) ' Find continuous PermitD's of same rate with quantity can cover bOH(J).
'                                                                          ' After found, return mAyPermitD, Rate, oIdx and set QtyAy to zero
If Sz(PermitDAy) > 0 Then
    ZRun "Update PermitD set BchNo='?' where PermitD in (?) and SKU='?'", BchNo, AyJnComma(PermitDAy), Sku
    ZRun "Insert into SkuB (Sku,BchNo,DutyRateB) values ('?','?',?)", Sku, BchNo, Rate
End If
End Sub

Private Function RR_2BchCurrentDb_3AyPermitD(OH&, PermitD&(), Rate@(), Qty&(), ByRef ORate@) As Long()
'Aim: Find continuous PermitD's of same rate with quantity can cover OH(J).
'     After found, return mAyPermitD, oRate, oIdx
Dim Idx&()
Dim U&
    Idx = RR_2BchCurrentDb_3AyPermitD_1AyIdx(OH, Rate, Qty)
    U = UB(Idx&): If U < 0 Then Exit Function
ORate = Rate(Idx&(0))

Dim O&()
    ReDim O(U)
    Dim I%
    For I = 0 To UBound(Idx)
        O(I) = PermitD(Idx(I))
    Next
    RR_2BchCurrentDb_3AyPermitD = O
End Function

Private Function RR_2BchCurrentDb_3AyPermitD_1AyIdx(OH&, Rate@(), Qty&()) As Long()
'Aim: Find continuous records of same rate with quantity can cover bOH(J).
'     After found, return oAyIdx and set aQty to zero

Dim Idx&
    Idx = RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx(OH, Rate, Qty) ' Find oIdx from which onward, the pQty can cover pOH and having same rate.
If Idx = -1 Then Exit Function
Dim Rate_@: Rate_ = Rate(Idx)
Dim O&()
Dim J&
Dim OH_&: OH_ = OH
For J = Idx To UBound(Qty)
    If Rate(J) <> Rate_ Then RR_2BchCurrentDb_3AyPermitD_1AyIdx = O: Exit Function
    OH_ = OH_ - Qty(J)
    If Qty(J) > 0 Then Push O, J
    If OH_ <= 0 Then RR_2BchCurrentDb_3AyPermitD_1AyIdx = O: Exit Function
Next
Stop ' impossible to reach here.
End Function

Private Function RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx&(OH&, Rate@(), Qty&())
'Aim: ' Find OIdx from which onward, the oQty can cover {OH} and having same rate.
Dim O&
    O = RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt(Qty, Rate)
While O <> -1
    Dim Rate1@: Rate1 = Rate(O)
    If RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_2IsOK(OH, O, Rate, Qty) Then
        RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx = O: Exit Function
    Else
        O = RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt(Qty, Rate, O + 1, Rate1)
    End If
Wend
RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx = -1
End Function

Private Function RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt%(QtyAy&(), RateAy@(), Optional BIdx&, Optional Rate@)
'Aim: Start from pBeg, find oIdx so that QtyAy(oIdx)>0 and pRate(oIdx)<>p_Rate
Dim O%
For O = BIdx To UBound(QtyAy)
    If QtyAy(O) <> 0 Then
        If Rate <> RateAy(O) Then RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt = O: Exit Function
    End If
Next
RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_1Nxt = -1
End Function

Private Function RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_2IsOK(OH&, Idx&, RateAy@(), QtyAy&()) As Boolean
'Return true if Idx is OKIdx:
'Use Idx & RateAy() to find Rate1
'From Idx onward, QtyAy of records of same Rate can cover OH return true else false
Dim Rate1@
Dim OH1&
    Rate1 = RateAy(Idx)
    OH1 = OH
Dim J&
For J = Idx To UBound(RateAy)
    If RateAy(J) <> Rate1 Then Exit Function
    If QtyAy(J) >= OH1 Then RR_2BchCurrentDb_3AyPermitD_1AyIdx_1Idx_2IsOK = True: Exit Function
    OH1 = OH1 - QtyAy(J)
Next
End Function

Private Sub Tmp1AssignOH_Bld()
'Aim: Create #Assign_OH from table-OH for those Tax-Paid item with SKU+BchNo not found in table PermitD
'Ref  #Assign_OH  = SKU,BchNo,OH
'     PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
Dim D As YMD: D = TblOHMaxYMD

TblDrp "#Assign_OH"
ZRun "Select Distinct SKU,BchNo,Sum(Bott) as OH" & _
" into `#Assign_OH` from OH x" & _
" where YY=? and MM=? and DD=?" & _
" and YpStk in (Select YpStk from YpStk where IsTaxPaid)" & _
" and SKU in (Select SKU from SKU_StkHld where TaxRate is not null)" & _
" group by SKU,BchNo" & _
" having Sum(Bott)<>0" & _
" order by BchNo desc", D.Y, D.M, D.D

ZRun "Update`#Assign_OH` x inner join PermitD a on a.Sku=x.Sku and a.BchNo=x.BchNo set x.OH=Null"
ZRun "Delete from `#Assign_OH` where OH is null"
End Sub

Private Sub Tmp2AssignSKU_Bld()
ZRun "Select Distinct SKU into `#Assign_SKU` from `#Assign_OH`"
End Sub

Private Sub Tmp3AssignLot_Crt()
ZRun "Create Table `#Assign_Lot` (Lot Integer, Sku Text(15), Rate Currency, MinPermitDate date, Qty Long," & _
" Constraint PrimaryKey Primary Key (Lot), Constraint `#Assign_Lot` Unique (Sku,Rate,MinPermitDate)) "
End Sub

Private Sub Tmp4AssignLotD_Crt()
ZRun "Create Table `?` (Lot Integer, PermitD Long, Constraint `#Assign_LotD` unique (Lot,PermitD))", C_TmpAssignLotD
End Sub

Private Function Upd_and_Ins_SqlAy(Sku) As Variant()
If Sku = "1034125" Then Stop
'Aim: Read *OH & *Lot
'     For each *OH find one or more *Lot covering the quantity
'     Set PermitD->BchNo & Write to SkuB
'Ref:#Assign_OH  = SKU,BchNo,OH
'    #Assign_Lot = Lot SKU Rate MinPermitDate | Qty
'    #Assign_LotD= Lot PermitD |
'    PermitD = Permit Sku | SeqNo Qty BchNo Rate Amt DteCrt DteUpd
Dim DrAy()
'----------------

Dim LotAy%(), RateAy@(), QtyAy&()
    DrAy = SqlDrAy(FmtQQ("Select Lot,Rate,Qty from `#Assign_Lot` where SKU='?' order by MinPermitDate Desc, Qty Desc", Sku))
    DrAyAsg DrAy, LotAy, RateAy, QtyAy

Dim OUpd_TblPermitD$()
Dim OIns_TblSkuB$()
    Dim Idx&
    Dim PermitDAy&()
    DrAy = SqlDrAy(FmtQQ("Select BchNo,OH from `#Assign_OH` where SKU='?' order by OH Desc", Sku))
    If AyIsEmpty(DrAy) Then Exit Function
    Dim Dr, BchNo$, OH&
    For Each Dr In DrAy
        AyAsg Dr, BchNo, OH
        Idx = AyCoverIdx(QtyAy, OH) ' Find those Lot's quantity can cover OHAy(J)
        If Idx >= 0 Then
            PermitDAy = SqlLngAy("Select PermitD from `#Assign_LotD` where Lot=" & LotAy(Idx))
            Push OUpd_TblPermitD, FmtQQ("Update PermitD set BchNo='?' where PermitD in (?) and SKU='?'", BchNo, AyJnComma(PermitDAy), Sku)
            Push OIns_TblSkuB, FmtQQ("Insert into SkuB (Sku,BchNo,DutyRateB) values ('?','?',?)", Sku, BchNo, RateAy(Idx))
        End If
    Next
Upd_and_Ins_SqlAy = Array(OUpd_TblPermitD, OIns_TblSkuB)
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

Tmp1AssignOH_Bld
Tmp2AssignSKU_Bld
Tmp3AssignLot_Crt
Tmp4AssignLotD_Crt

Dim mLot%: mLot = 1
Dim mLasSku$
Dim mLasRate$
Dim mMinPermitDate As Date
Dim mQty&
Dim mRsLot As Recordset:  Set mRsLot = CurrentDb.TableDefs(C_TmpAssignLot).OpenRecordset
Dim mRsLotD As Recordset: Set mRsLotD = CurrentDb.TableDefs(C_TmpAssignLotD).OpenRecordset
Dim mRs As Recordset:     Set mRs = CurrentDb.OpenRecordset("Select x.PermitD, Sku, Round(x.Rate,2) as Rate, PermitDate, x.Qty from PermitD x inner join Permit a on a.Permit=x.Permit order by Sku,PermitDate,Rate")
With mRs
    mRsLotD.AddNew
    mRsLotD!Lot = mLot
    mRsLotD!PermitD = !PermitD
    mRsLotD.Update          '<== Insert LotD
    
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
        mRsLotD.Update          '<===
        
        .MoveNext
    Wend
    mRsLot.AddNew
    mRsLot!Lot = mLot
    mRsLot!Sku = mLasSku
    mRsLot!Rate = mLasRate
    mRsLot!MinPermitDate = mMinPermitDate
    mRsLot!Qty = mQty
    mRsLot.Update     '<===
End With
End Sub

Private Sub zBldRecordSourceTable(xYY As Byte, xMM As Byte, xDD As Byte)
'Aim: Refresh table-(frmBchRate frmBchRateOH) from OH (Co=8600)
'Refresh frmBchRate    = Sku DesSku OH IsNoAssign DutyRate BottPerAc DutyRateBott
'Refresh frmBchRateOH = Sku BchNo OH IsNoAssign
'---- refresh table-frmBchRateOH ------------------------------------------
ZRun "Delete from frmBchRateOH"

'Inp: OH
'     YpStk
'     SKU_StkHld
ZRun "Insert into frmBchRateOH (Sku,BchNo,OH,IsNoAssign)" & _
" Select Distinct SKU, BchNo, Sum(Bott), False" & _
" from OH x inner join YpStk a on x.YpStk=a.YpStk" & _
" where IsTaxPaid and YY=? and MM=? and DD=? and Co=1" & _
" and SKU in (Select SKU from SKU_StkHld where TaxRate is not null)" & _
" group by SKU,BchNo" & _
" having Sum(x.bott)<>0", xYY, xMM, xDD    ' 'Co=1 is 8600 HK

ZRun "Select Sku,BchNo,Sum(Qty) as QtyPermit into `#frmBchRate_Permit` from PermitD where Nz(BchNo,'')<>'' Group by Sku,BchNo"
ZRun "Update frmBchRateOH x left join `#frmBchRate_Permit` a on a.Sku=x.Sku and a.BchNo=x.BchNo set IsNoAssign=True where OH>Nz(QtyPermit,0)"
TblDrp "#frmBchRate_Permit"

'---- refresh table-frmBchRate ------------------------------------------
'Inp: table-frmBchRateOH
'     table-Sku_StkHld
'     table-tblSKU
'     table-Sku88
'     query-qSku
'Oup: frmBchRate
ZRun "Delete from frmBchRate"
ZRun "Insert into frmBchRate (Sku,OH,IsNoAssign) Select Distinct Sku, Sum(OH), Min(IsNoAssign) from frmBchRateOH group by Sku"
ZRun "Update frmBchRate x inner join Sku_StkHld a on a.Sku=x.Sku set x.DutyRate=a.TaxRate"
ZRun "Update frmBchRate x inner join tblSKU a on a.SkuTxt=x.Sku set x.BottPerAc=a.BtlPerCs"
ZRun "Update (frmBchRate x inner join Sku88 a on a.Sku=x.Sku) inner join tblSku b on b.SkuTxt=a.Sku88 set x.BottPerAc=b.BtlPerCs"
ZRun "Update frmBchRate set DutyRateBott=DutyRate/BottPerAc where Nz(BottPerAc,0)<>0"
ZRun "Update frmBchRate x inner join qSku a on a.Sku=x.Sku set DesSku=`Sku Description`"
End Sub

Private Sub zBldRecordSourceTable__Tst()
zBldRecordSourceTable 17, 3, 31
End Sub

Private Sub ZOpn(T$)
DoCmd.OpenTable T
End Sub

Private Sub ZRun(Sql$, ParamArray Ap())
Dim Av(): Av = Ap
Dim S$: S = FmtQQAv(Sql, Av)
DoCmd.RunSql S
End Sub

Private Sub ZRunAy(SqlAy$())
Dim I
For Each I In SqlAy
    DoCmd.RunSql I
Next
End Sub

