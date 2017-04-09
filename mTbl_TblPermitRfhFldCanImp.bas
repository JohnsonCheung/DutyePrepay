Attribute VB_Name = "mTbl_TblPermitRfhFldCanImp"
Option Compare Database
Option Explicit

Sub TblPermitRfhFldCanImp()
Dim A1$(): A1 = ImpPermitNoAy
Dim A2$(): A2 = S_PermitNoAy2
Dim B$():   B = AyMinus(A1, A2)
                S_InsPermitAy B
                S_Upd A1
End Sub

Private Sub S_InsPermit(PermitNo$, Dte$, Ac$, AcNm$, Bnk$, Usr$)
Dim Pst$
Dim Crt$
    Pst = Dte
    Crt = Now
SqlRun FmtQQ("Insert into Permit (PermitNo,PermitDate,PostDate,GLAc,GLAcName,BankCode,ByUsr,DteCrt,IsCur,CanImp) values ('?',#?#,#?#,'?','?','?','?',#?#,False,True)", _
PermitNo, Dte, Pst, Ac, AcNm, Bnk, Usr, Crt)
End Sub

Private Sub S_InsPermitAy(PermitNoAy$())
If AyIsEmpty(PermitNoAy) Then Exit Sub
Dim Dte$
    Dte = Format(Date, "YYYY-MM-DD")
Dim Ac$
Dim AcNm$
Dim Bnk$
Dim Usr$:
SqlAsg "Select GLAc,GLAcName,BankCode,ByUsr from Default", Ac, AcNm, Bnk, Usr
Dim J%
For J = 0 To UB(PermitNoAy)
    S_InsPermit PermitNoAy(J), Dte, Ac, AcNm, Bnk, Usr
Next
End Sub

Private Sub S_InsPermitAy__Tst()
Dim A$(): A = Split("a b c")
S_InsPermitAy A
End Sub

Private Function S_PermitNoAy2() As String()
S_PermitNoAy2 = SqlSy("Select PermitNo from Permit")
End Function

Private Sub S_PermitNoAy2__Tst()
AyBrw S_PermitNoAy2
End Sub

Private Sub S_Upd(PermitNo$())
SqlRun "Update Permit Set CanImp=False"
If AyIsEmpty(PermitNo) Then Exit Sub
Dim A$: A = Join(AyQuote(PermitNo, "'"), ",")
SqlRun FmtQQ("Update Permit Set CanImp=True where PermitNo in (?)", A)
End Sub

