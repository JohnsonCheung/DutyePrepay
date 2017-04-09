Attribute VB_Name = "mCmd_FrmPermitCmdExpList"
Option Compare Database
Option Explicit

Sub FrmPermitCmdExpList()
Dim P$: P = FbCurPth
Dim Fx$: Fx = P & "Output\Permit Report.xlsx"
Dim Tp$: Tp = P & "Template\Template_PermitRpt.xlsx"
SqlRun "SELECT SnoFinStream,SnoSHBrandG,SnoHse,SnoSHBrand,SnoB,NmFinStream,CdSHBrandG,NmHse,NmSHBrand,NmB,x.PermitNo,Year(x.PermitDate) as PermitYear, Month(x.PermitDate) as PermitMonth, Day(x.PermitDate) as PermitDay, x.PermitDate, x.PostDate, x.GLAc, x.GLAcName, x.BankCode, x.ByUsr, x.DteCrt AS DteCrtHdr, x.DteUpd AS DteUpdHdr, a.Sku, b.DesSku, a.SeqNo, a.Qty, a.BchNo, a.Rate, a.Amt, a.DteCrt, a.DteUpd" & _
" Into `@PermitRpt`" & _
" FROM (Permit AS x INNER JOIN PermitD AS a ON x.Permit = a.Permit) Left Join q2Sku b on b.Sku=a.Sku" & _
" Order by SnoFinStream,SnoSHBrandG,CdSHBrandG,SnoHse,NmHse,SnoSHBrand,NmSHBrand,SnoB,NmB,a.SKU,PermitDate Desc"
FxCrtFmTpWithRfh Fx, Tp, Vis:=True
'== Reseq ====================
End Sub
