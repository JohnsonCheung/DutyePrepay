Attribute VB_Name = "nIde_nConstSrcLin_OptExplicit"
Option Compare Database
Option Explicit

Sub OptExplicitEnsMd(Optional A As CodeModule)
Dim B As CodeModule: Set B = MdNz(A)
Dim C$(): C = MdDclLy(B)
Debug.Print MdNm(B);
Dim J%
For J = 0 To UB(C)
    If C(J) = "Option Explicit" Then
        Debug.Print
        Exit Sub
    End If
Next
Debug.Print "<=== [Option Explicit] Inserted"
B.InsertLines 2, "Option Explicit"
End Sub

Sub OptExplicitEnsPj(Optional A As vbproject)
Dim MdAy() As CodeModule: MdAy = PjMdAy(A)
AyEachEle MdAy, "OptExplicitEnsMd"
End Sub
