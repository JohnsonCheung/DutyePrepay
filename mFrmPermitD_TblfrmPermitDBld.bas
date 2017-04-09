Attribute VB_Name = "mFrmPermitD_TblfrmPermitDBld"
Option Compare Database
Option Explicit

Sub TblfrmPermitDBld(Permit&)
'Aim: Build table-frmPermitD by Permit
SqlRun "Delete from frmPermitD"
SqlRun "Insert into frmPermitD (PermitD,DesSku) select PermitD,Nz(`SKU Description`,'') from PermitD x left join qSKU a on a.Sku=x.Sku where Permit=" & Permit
SqlRun "Update ((frmPermitD x inner join PermitD a on a.PermitD=x.PermitD) inner join Sku_StkHld b on b.Sku=a.Sku) inner join qSKU c on c.Sku=a.Sku set x.DutyRateZHT0=IIf(Nz(BtlPerCs,0)=0,0,b.TaxRate/BtlPerCs)"
End Sub

Sub TblfrmPermitDBld__Tst()
TblfrmPermitDBld 1952
End Sub
