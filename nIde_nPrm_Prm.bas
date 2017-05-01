Attribute VB_Name = "nIde_nPrm_Prm"
Option Compare Database
Option Explicit

Function PrmAy(PrmStr$) As Prm()
Dim S$: S = Trim(PrmStr)
If S = "" Then Exit Function
Dim Ay$(): Ay = Split(PrmStr, ",")
Dim O() As Prm
Dim J%
Dim U%: U = UB(Ay)
ReDim O(U)
For J = 0 To U
    Set O(J) = Prm(Ay(J))
Next
PrmAy = O
End Function

Function PrmAy__Tst()
Dim PrmStr$
Dim Act() As Prm
Dim mCase As Byte
mCase = 2
Select Case mCase
Case 1: PrmStr = "ByVal pPrmDcl$"
Case 2: PrmStr = "Optional ByVal pPrmDcl As String = ""ABC"""
Case 3: PrmStr = "pPrmDcl$()"
End Select
Act = PrmAy(PrmStr)
Debug.Assert PrmAyToStr(Act) = PrmStr
End Function

Function PrmAyToStr$(A() As Prm)
PrmAyToStr = JnComma(AyMapIntoSy(A, "PrmToStr"))
End Function

Function PrmToStr$(A As Prm)
With A
Dim Opt$: If .IsOpt Then Opt = "Optional "
Dim ByV$: If .IsByVal Then ByV = "ByVal "
Dim PrmAy$: If .IsPrmAy Then PrmAy = "Paramarray "
Dim Nm$:  Nm = .Nm
Dim TyChr$: TyChr = .TyChr
Dim Bkt$: If .IsAy Then Bkt = "()"
Dim AsTy$: If .AsTy <> "" Then AsTy = " As " & .AsTy
Dim Dft$: If .DftStr <> "" Then Dft = " = " & .DftStr
End With
PrmToStr = Opt & ByV & Nm & TyChr & Bkt & AsTy & Dft
End Function

Private Function Prm(OnePrmStr) As Prm
Dim O As New Prm
Dim L$: L = OnePrmStr
With O
    If ParseStr(L, "Optional ") Then .IsOpt = True
    If ParseStr(L, "ByVal ") Then
        .IsByVal = True
    ElseIf ParseStr(L, "ByRef ") Then
        .IsByVal = False
    ElseIf ParseStr(L, "Paramarray ") Then
        .IsPrmAy = True
    End If
    .Nm = ParseNm(L)
    .TyChr = ParseChr(L, "")
    If ParseStr(L, "()") Then .IsAy = True
    If ParseStr(L, " As ") Then
        .AsTy = ParseNm(L)
    End If
    If ParseStr(L, " = ") Then
        .DftStr = Trim(L)
    End If
End With
Set Prm = O
End Function
