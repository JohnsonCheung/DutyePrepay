Attribute VB_Name = "nDte_YYYYMMDD"
Option Compare Database
Option Explicit
Type tYYYYMM
    YYYY As Integer
    MM As Byte
End Type

Function Cv_YYYYMM2FirstDte(P As tYYYYMM) As Date
With P
    Cv_YYYYMM2FirstDte = CDate(.YYYY & "/" & .MM & "/1")
End With
End Function

Function Cv_YYYYMM2LasDte(P As tYYYYMM) As Date
With Cv_YYYYMM2Nxt(P)
    Cv_YYYYMM2LasDte = CDate(.YYYY & "/" & .MM & "/1") - 1
End With
End Function

Function Cv_YYYYMM2Nxt(pYYYYMM As tYYYYMM) As tYYYYMM
Dim mYYYYMM As tYYYYMM
With pYYYYMM
    If .MM = 12 Then
        mYYYYMM.MM = 1
        mYYYYMM.YYYY = .YYYY + 1
    Else
        mYYYYMM.MM = .MM + 1
        mYYYYMM.YYYY = .YYYY
    End If
End With
Cv_YYYYMM2Nxt = mYYYYMM
End Function

Function Cv_YYYYMM2Prv(pYYYYMM As tYYYYMM) As tYYYYMM
Dim mYYYYMM As tYYYYMM
With pYYYYMM
    If .MM = 1 Then
        mYYYYMM.MM = 12
        mYYYYMM.YYYY = .YYYY - 1
    Else
        mYYYYMM.MM = .MM - 1
        mYYYYMM.YYYY = .YYYY
    End If
End With
Cv_YYYYMM2Prv = mYYYYMM
End Function

Function Dte2YYYYMM(pDte As Date) As tYYYYMM
Dim mYYYYMM As tYYYYMM
With mYYYYMM
    .YYYY = Year(pDte)
    .MM = Month(pDte)
End With
Dte2YYYYMM = mYYYYMM
End Function
