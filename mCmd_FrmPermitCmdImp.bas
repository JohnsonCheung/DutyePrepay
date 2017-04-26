Attribute VB_Name = "mCmd_FrmPermitCmdImp"
Option Compare Database
Option Explicit

Sub FrmPermitCmdImp(PermitNo$)
                          ZLnk PermitNo
Dim Er():            Er = ZChk
                          If AyHasEle(Er) Then ErBrw ErApd(Er, "ImpErr"): Exit Sub
Dim W$:               W = FmtQQ("PermitNo='?'", PermitNo)
Dim PermitId&: PermitId = TblFldToLng("Permit", "Permit", W)
                          ZInsPermitD PermitId
                          ZUpdPermit PermitNo
                          ZMovFil PermitNo
                          TblDrp ">Permit"
End Sub

Sub FrmPermitCmdImp__Tst()
FrmPermitCmdImp ImpPermitFstNo
End Sub

Private Function ZChk() As Variant()
ZChk = AyAdd(ZChk_FldTy, ZChk_EmptyRec)
End Function

Private Function ZChk_EmptyRec() As Variant()
ZChk_EmptyRec = TblChkEmptyRec(">Permit")
End Function

Private Function ZChk_FldTy() As Variant()
Const C = "TXT : [Batch Number] | DBL : SKU [Order Qty#]"
'Const C = _
'"TXT : [Material Description] Warehouse [Age Certificate] [Com# Code] [Permit Number] [Batch Number] [Customer / Duty Paid] | " & _
'"DBL : Plant SKU [Order Qty#] [No of case] | " & _
'"DTE : Remark"
ZChk_FldTy = TblFldChkFnStr(">Permit", C)
End Function

Private Sub ZInsPermitD(PermitId&)
Dim S As Dt:       S = ImpPermitDt
Const F$ = "Permit SKU SeqNo Qty BchNo"
Dim Sku$, Qty&, BchNo$, SeqNo%
Dim J&, Dr()
SeqNo = 0
SqlRun FmtQQ("Delete From PERMITD WHERE Permit=?", PermitId)
For J = 0 To UB(S.DrAy)
    Dr = S.DrAy(J)
    AyAsg Dr, Sku, BchNo, Qty
    SeqNo = SeqNo + 10
    SqlRun SqsOfIns("PermitD", F, ApAv(PermitId, Sku, SeqNo, Qty, BchNo))
Next
End Sub

Private Sub ZInsPermitD__Tst()
ZInsPermitD ImpPermitFstNo
End Sub

Private Sub ZLnk(PermitNo$)
LnkCrt_Fx ">Permit", ImpPermitFdr & PermitNo & ".xlsx"
End Sub

Private Sub ZLnk__Tst()
ZLnk ImpPermitFstNo
TblBrw ">Permit"
End Sub

Private Sub ZMovFil(PermitNo$)
Dim DirFm$: DirFm = ImpPermitFdr
Dim DirTo$:
    DirTo = DirFm & "Done\":                                PthEns DirTo
    DirTo = DirTo & Format(Now, "YYYY-MM-DD hhmmss") & "\": PthEns DirTo
FfnMov DirFm & PermitNo & ".xlsx", DirTo
End Sub

Private Sub ZUpdPermit(PermitNo$)
Dim Permit&: Permit = TblPermitIdByNo(PermitNo)
Dim Q&
Dim N%
    Dim S1$: S1 = FmtQQ("SELECT SUM(x.QTY) as QTY, COUNT(*) as N FROM PERMITD x WHERE PERMIT=?", Permit)
    SqlAsg S1, Q, N
Dim S$: S = FmtQQ("UPDATE Permit Set Qty=?,NSku=?,CanImp=false,DteImp=Now() where Permit=?", Q, N, Permit)
SqlRun S
End Sub

Private Sub ZUpdPermit__Tst()
ZUpdPermit ImpPermitFstNo
End Sub

