Attribute VB_Name = "mInf_ImpPermit"
Option Compare Database
Option Explicit

Function ImpPermitDt() As Dt
Dim N As Date:     N = Now
Const T$ = ">Permit"
Const F$ = "CSTR(X.SKU) as SKU, [Batch Number] as BchNo, [Order Qty#] as Qty"
Dim S$: S = SqsOfSel(T, F)
ImpPermitDt = SqlDt(S)
End Function

Sub ImpPermitDt__Tst()
DtBrw ImpPermitDt
End Sub

Function ImpPermitFdr$()
Dim O$: O = FbCurPth & "SAPDownloadExcel\Permit\"
PthEns O
ImpPermitFdr = O
End Function

Function ImpPermitFstFxFn$()
Dim A$(): A = ImpPermitFxFnAy: If AyIsEmpty(A) Then Exit Function
ImpPermitFstFxFn = A(0)
End Function

Function ImpPermitFstNo$()
Dim A$: A = ImpPermitFstFxFn: If A = "" Then Exit Function
ImpPermitFstNo = Left(A, Len(A) - 5)
End Function

Function ImpPermitFxFnAy() As String()
' All *.xlsx Fnn in ImpPermitFdr
ImpPermitFxFnAy = PthFnAy(ImpPermitFdr, "*.xlsx")
End Function

Function ImpPermitNoAy() As String()
' All *.xlsx Fnn in ImpPermitFdr
ImpPermitNoAy = PthFnnAy(ImpPermitFdr, "*.xlsx")
End Function

Private Sub ImpPermitNoAy__Tst()
AyBrw ImpPermitNoAy
End Sub
