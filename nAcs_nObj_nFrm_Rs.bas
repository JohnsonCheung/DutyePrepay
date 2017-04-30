Attribute VB_Name = "nAcs_nObj_nFrm_Rs"
Option Compare Database
Option Explicit

Sub RsPutFrm(A As DAO.Recordset, NmStr$, Frm As Access.Form)
'Aim: Copy the fields value from {Rs} to the controls in {pFrm}.  Only those fields in {FnStr} will be copied.
'     {FnStr} is in fmt of aaa=xxx,bbb,ccc  aaa,bbb,ccc will be field name in {pFrm} & xxx,bbb,ccc will be field in {pRs}
Dim FrmNy$(), RsNy$(), GivenFny$()
Dim Ny$()
    Ny = AyIntersect(FrmNy, RsNy)
Dim J%, RsV, FrmV
For J = 0 To UB(Ny)
    With Frm.Controls(Ny(J))
        RsV = A.Fields(RsNy(J)).Value
        FrmV = .Value
        If RsV <> FrmV Then .Value = RsV  '<------
    End With
Next
FrmSavRec Frm
End Sub
