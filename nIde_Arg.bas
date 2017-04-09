Attribute VB_Name = "nIde_Arg"
Option Compare Database
Option Explicit
Type d_Arg
x_NmArg As String
x_NmTypArg As String
x_IsAy As Boolean
x_IsPrmAy As Boolean
x_IsOpt As Boolean
x_IsByVal As Boolean
x_DftVal As Variant
End Type

Function ArgBrk(oDArg As d_Arg, pArgDcl$) As Boolean
'Aim: Brk pPrcBody in fmt (...) into oAyDArg()
Const cSub$ = "BrkArgDcl"
'    Public x_IsAs Boolean, x_NmArg$, x_NmTypArg$, x_IsOpt As Boolean, x_DftVal
Dim mArgDcl$: mArgDcl = Trim(Replace(pArgDcl, "_" & vbCrLf, ""))
With oDArg
    Dim mP%, mA$
    .x_IsOpt = False
    .x_IsByVal = False
    .x_IsAy = False
    .x_IsPrmAy = False
    mA = "Optional ":   mP = InStr(mArgDcl, mA): .x_IsOpt = (mP > 0):    If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))
    mA = "ByVal ":      mP = InStr(mArgDcl, mA): .x_IsByVal = (mP > 0):  If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))
    mA = "ByRef ":      mP = InStr(mArgDcl, mA):                         If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))
    mA = "ParamArray ": mP = InStr(mArgDcl, mA): .x_IsPrmAy = (mP > 0):  If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))
    mA = "()":          mP = InStr(mArgDcl, mA): .x_IsAy = (mP > 0):     If mP > 0 Then mArgDcl = Trim(Replace(mArgDcl, mA, ""))

    mP = InStr(mArgDcl, " = ")
    If mP > 0 Then
        .x_DftVal = Trim(Mid(mArgDcl, mP + 3))
        mArgDcl = Trim(Left(mArgDcl, mP - 1))
    Else
        .x_DftVal = ""
    End If

    mP = InStr(mArgDcl, " As ")
    If mP > 0 Then
        .x_NmTypArg = Trim(Mid(mArgDcl, mP + 3))
        Select Case .x_NmTypArg
        Case "String": .x_NmTypArg = "$"
        Case "Long": .x_NmTypArg = "&"
        Case "Integer": .x_NmTypArg = "%"
        Case "Single": .x_NmTypArg = "!"
        Case "Double": .x_NmTypArg = "#"
        Case "Currency": .x_NmTypArg = "@"
        End Select
        .x_NmArg = Trim(Left(mArgDcl, mP - 1))
    Else
        Dim mB$
        mB = Right(mArgDcl, 1)
        Select Case mB
        Case "%", "$", "&", "#", "!":   .x_NmTypArg = mB:           .x_NmArg = Left(mArgDcl, Len(mArgDcl) - 1)
        Case Else:                      .x_NmTypArg = "Variant":    .x_NmArg = mArgDcl
        End Select
    End If
End With
Exit Function
E: ArgBrk = True: ss.B cSub, cMod, "pArgDcl", pArgDcl
End Function

Function ArgBrk__Tst()
Dim mArgDcl$
Dim mDArg As d_Arg
Dim mCase As Byte
mCase = 2
Select Case mCase
Case 1: mArgDcl = "ByVal pArgDcl$"
Case 2: mArgDcl = "Optional ByVal pArgDcl As String = ""ABC"""
Case 3: mArgDcl = "pArgDcl$()"
End Select
If ArgBrk(mDArg, mArgDcl) Then Stop: GoTo E
Shw_DbgWin

Debug.Print Fct.UnderlineStr(mArgDcl, "*")
Debug.Print mArgDcl
Debug.Print Fct.UnderlineStr(mArgDcl)

'Debug.Print ToStr_DArg(mDArg, False)
Exit Function
E:
End Function

Function ArgBrkPrmDcl(oAyDArg() As d_Arg, pPrmDcl$) As Boolean
'Aim: Brk pPrcBody in fmt (...) into oAyDArg()
Const cSub$ = "BrkArg"
If pPrmDcl = "" Then Erase oAyDArg: Exit Function
Dim mAyArgDcl$(): mAyArgDcl = Split(Replace(Replace(pPrmDcl, "_" & vbCrLf, " "), vbCrLf, " "), CtComma)
Dim J%, N%: N% = Siz_Ay(mAyArgDcl)
ReDim oAyDArg(N - 1)
For J = 0 To N - 1
    If ArgBrk(oAyDArg(J), mAyArgDcl(J)) Then ss.A 1: GoTo E
Next
Exit Function
E: ArgBrkPrmDcl = True: ss.B cSub, cMod, "pPrmDcl", pPrmDcl
End Function

Sub ArgBrkPrmDcl__Tst()
Dim mPrmDcl$, mAyDArg() As d_Arg
Dim mPrcBody$, mNmPrc$, mNmPrj_Nmm$
Dim mCase As Byte
mCase = 7
Select Case mCase
Case 1: mPrmDcl = "oDArg As d_Arg, pArgDcl$"
Case 2: mPrmDcl = "Optional pArgDcl As String = ""ABC"""
Case 3: mPrmDcl = "Optional pInclTbl As Boolean = True, Optional ByVal pInclQry As Boolean = True, Optional pInclTypFld As Boolean = False, Optional pCls As Boolean = False"
Case 4: mNmPrj_Nmm = "Bld":      mNmPrc = "Lst_ByAyV":
Case 5: mNmPrj_Nmm = "ToStr":    mNmPrc = "Ays"
Case 6: mNmPrj_Nmm = "ToStr":    mNmPrc = "Ays"
Case 7: mNmPrj_Nmm = "Run":      mNmPrc = "qBrkRec"
End Select
If 4 <= mCase And mCase <= 7 Then
    If Fnd_PrcBody(mPrcBody, mNmPrj_Nmm, mNmPrc) Then Stop: GoTo E
    mPrmDcl = Cut_Prm(mPrcBody)
End If
If ArgBrkPrmDcl(mAyDArg, mPrmDcl) Then Stop: GoTo E
Shw_DbgWin

Debug.Print Fct.UnderlineStr(mPrmDcl, "*")
Debug.Print mPrmDcl
Debug.Print Fct.UnderlineStr(mPrmDcl)
Dim J%
For J = 0 To UBound(mAyDArg)
    Stop
   ' Debug.Print ToStr_DArg(mAyDArg(J))
Next
Exit Sub
E:
End Sub
