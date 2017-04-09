Attribute VB_Name = "nIde_nPos_Mth"
Option Compare Database
Option Explicit

Function MthLCC(MthNm$, Optional A As CodeModule) As LCC
Dim Md As CodeModule: Set Md = MdNz(A)
Dim O As LCC
    Dim Lin$
    With O
        .L = MthLno(MthNm, A)
        If .L > 0 Then
            Lin = Md.Lines(.L, 1)
            .C1 = InStr(Lin, MthNm)
            .C2 = .C1 + Len(MthNm)
        End If
    End With
MthLCC = O
End Function

Sub MthLCC__Tst()
Dim Act$
Act = LCCToStr(MthLCC("MthLCC")): Debug.Assert Act = "L8 C(10 16)"
End Sub

Function MthLCC_ForEdt(MthNm$, Optional A As CodeModule) As LCC
Dim B As LCC: B = MthLCC(MthNm, A)
If B.L = 0 Then Exit Function
Dim O As LCC
With O
    .L = B.L + 1
    .C1 = 1
    .C2 = 2
End With
MthLCC_ForEdt = O
End Function
