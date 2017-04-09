Attribute VB_Name = "nDao_Cnn"
Option Compare Database
Option Explicit

Function CnnDt(Optional D As database) As Dt
Dim O As Dt
O.Fny = Split("TblNm AppNm Ver Ext Msg CnnStr")
Dim A$()
    A = CnnSy(D)
Dim U%
    U = UB(A)
If U = -1 Then Exit Function
Dim R()
    ReDim R(U)
    Dim J%
    For J = 0 To UB(A)
        R(J) = CnnStrDr(A(J))
    Next
O.DrAy = R
CnnDt = DtSrt(O, "AppNm- TblNm-")
End Function

Sub CnnDt__Tst()
DtBrw CnnDt, "AppNm"
End Sub

Sub CnnDtBrw(Optional D As database)
DtBrw CnnDt(D)
End Sub

Function CnnSy(Optional D As database) As String()
Dim T As DAO.TableDef
Dim O$()
For Each T In DbNz(D).TableDefs
    If T.Connect <> "" Then
        Push O, T.Name & "|" & T.Connect
   End If
Next
CnnSy = O
End Function

Sub CnnSy__Tst()
AyBrw CnnSy
End Sub
