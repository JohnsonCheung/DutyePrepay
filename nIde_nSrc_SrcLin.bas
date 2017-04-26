Attribute VB_Name = "nIde_nSrc_SrcLin"
Option Compare Database
Option Explicit

Function SrcLinEnumNm$(SrcLin)
Dim A$: A = SrcLin
ParseSy A, ApSy("Private ", "Public ")
If ParseStr(A, "Enum ") = "" Then Exit Function
SrcLinEnumNm = A
End Function

Function SrcLinIsDim(SrcLin) As Boolean
'Debug.Print SrcLin
Dim O As Boolean
O = IsPfx(LTrim(SrcLin), "Dim ")
'If IsPfx(SrcLin, "Dim ") Then Stop
SrcLinIsDim = O
End Function

Sub SrcLinIsDim__Tst()
Debug.Assert SrcLinIsDim("Dim A") = True
End Sub

Function SrcLinIsMth(SrcLin) As Boolean
Dim A$: A = SrcLin
Dim B$: B = ParseSy(A, ApSy("Public ", "Private ", "Friend "))
Dim C$: C = ParseSy(A, ApSy("Function ", "Sub ", "Property ")): If C = "" Then Exit Function
SrcLinIsMth = True
End Function

Function SrcLinIsTth(Lin) As Boolean
If Not SrcLinIsMth(Lin) Then Exit Function
Dim A As MthBrk: A = MthBrkNew(Lin)
If A.Mdfy = "" Or A.Mdfy = "Public" Then
    If IsPfx(A.Nm, "Tst_") Then
        If A.Ty = "Sub" Then
            SrcLinIsTth = True
        Else
            Er "{Lin} has MthNm like [Tst_*] of not Ty=Sub", Lin
        End If
    End If
End If
End Function

Function SrcLinRmvRmk$(SrcLin)
SrcLinRmvRmk = StrBrk1(RmvStrTok(SrcLin), "'").S1
End Function

Function SrcLinTyNm$(SrcLin)
Dim A$: A = SrcLin
ParseSy A, ApSy("Private ", "Public ")
If ParseStr(A, "Type ") = "" Then Exit Function
SrcLinTyNm = A
End Function
