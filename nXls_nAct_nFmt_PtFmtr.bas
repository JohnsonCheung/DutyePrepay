Attribute VB_Name = "nXls_nAct_nFmt_PtFmtr"
Option Compare Database
Option Explicit
Type PtFmtr
    Fny() As String
    Row() As String
    Col() As String
    Dta() As String
    Pag() As String
    DtaSumFun() As XlConsolidationFunction
    DtaSumFld() As String
    DtaSumFmt() As String
    DtaSumFno() As Integer  ' The Field# (started from 1) of DtaFld within PT.DataFields
    LblVal()    As String
    LblFld()    As String   ' <LblFld> are <Fny> required to change the PivotField.Caption by <LblVal>
    LblDtaFno() As Integer  ' <LblDtaFno> are FieldNo in PivotTable.DataFields of those <LblFld> which is DataField.
                            '             To the change the Caption of a DataFields must use PivotTable.DataFields(<LblDtaFno>).Caption = <LblVal>
                            '             Using this will cause error                        PivotTable.PivoatFields(<LblFld>).Caption = <LblVal>
    LblColFld() As String   ' <LblColFld> is <LblFld> - <Dta>.  That means it is those non-Dta-Fld in <LblFld>
                            '             To the change the Caption of a Non-DataFields
                            '                Using this will be OK PivotTable.PivoatFields(<LblColFld>).Caption = <LblVal>
    SubTotFld() As String
    SubTotFno() As Integer
    WdtVal()    As Integer
    WdtFld()    As String
    WdtFno()    As Integer
    GrandColTot As Boolean
    GrandColWdt As Integer
    GrandRowTot As Boolean
    OutLinFld() As String
    OutLinFno() As Integer
    OutLinLvl() As Byte
    OpnInd      As Boolean
    Er()        As Variant
End Type

Function PtFmtrByFt(PtFmtrFt$) As PtFmtr
PtFmtrByFt = PtFmtr(FtLy(PtFmtrFt))
End Function

Function PtFmtr(PtFmtrLy$()) As PtFmtr
Dim L, A As S1S2, S2$
Dim Ly$(): Ly = AyExcl(PtFmtrLy, "StrIsBlank")
'-------------
Dim Fny$()
    For Each L In Ly
        A = StrBrk(L, ":")
        S2 = A.S2
        Select Case A.S1
        Case "Fny": Fny = Split(S2, " "): Exit For
        End Select
    Next

Dim O As PtFmtr
    If AyIsEmpty(Fny) Then
        O.Er = ErNew("no {Fny} is found in {PtFmtrLy}")
        PtFmtr = O
        Exit Function
    End If

    Dim Er()
    With O
        .Fny = Fny
        For Each L In Ly
            A = StrBrk(L, ":")
            S2 = A.S2
            Select Case A.S1
            Case "Fny":
            Case "Lbl": AyAsg ZBrk_Fmt_Lbl(S2, Fny, .Dta), _
                                                                            Er, .LblFld, .LblDtaFno, .LblColFld, .LblVal
            Case "Row": AyAsg ZBrkFld(S2, Fny, "Row", Sz(.Row) > 0), _
                                                                            Er, .Row
            Case "Col": AyAsg ZBrkFld(S2, Fny, "Col", Sz(.Col) > 0), _
                                                                            Er, .Col
            Case "Pag": AyAsg ZBrkFld(S2, Fny, "Pag", Sz(.Pag) > 0), _
                                                                            Er, .Pag
            Case "Dta": AyAsg ZBrkFld(S2, Fny, "Dta", Sz(.Dta) > 0), _
                                                                            Er, .Dta
            Case "Wdt": AyAsg ZBrk_Fmt_Wdt(S2, Fny), _
                                                                            Er, .WdtFld, .WdtFno, .WdtVal
            Case "OutLin": AyAsg ZBrk_Fmt_OutLin(S2, Fny), _
                                                                            Er, .OutLinFld, .OutLinFno, .OutLinLvl
            Case "DtaSum": AyAsg ZBrk_Tot_DtaSum(S2, .Dta), _
                                                                            Er, .DtaSumFld, .DtaSumFno, .DtaSumFun, .DtaSumFmt
            Case "SubTot": AyAsg ZBrkFld(S2, Fny, "SubTot", Sz(.SubTotFld) > 0), _
                                                                            Er, .SubTotFld
            Case "OpnInd": AyAsg ZBrkBool(S2, "OpnInd"), _
                                                                            Er, .OpnInd
            Case "GrandColTot": AyAsg ZBrk_Tot_GrandColTot(S2), _
                                                                            Er, .GrandColTot, .GrandColWdt
            Case "GrandRowTot": AyAsg ZBrkBool(S2, "GrandRowTot"), _
                                                                            Er, .GrandRowTot
            Case Else
                Er = ErNew("Lin [" & L & "] has invalid type.  Valid Type are [Lbl Row Col Pag Dta Fmt Wdt OutLin SubTot DtaSum GrandColTot GrandRowTot]")
            End Select
            PushAy O.Er, Er
        Next
    End With
    
    O.SubTotFno = ZFldIdx(O.SubTotFld, Fny)

PtFmtr = O
If Not AyIsEmpty(O.Er) Then
    PushAy O.Er, ErNew("Fny: " & Join(Fny, " "))
    PushAy O.Er, ErNew("DtaFld: " & Join(O.Dta, " "))
    PushAy O.Er, ErNew("There are errors (above) in given {PtFmtrLy} see below:")
    Dim J%
    For J = 0 To UB(PtFmtrLy)
        PushAy O.Er, ErNew(PtFmtrLy(J))
    Next
    ErBrw O.Er
End If
PtFmtr = O
End Function

Sub PtFmtr__Tst()
Dim A$()
Push A, "Row: AA BB CC X"
Push A, "Col: CC DD EE"
Push A, "Pag: DD EE"
Push A, "Dta: DD FF"
Push A, "GrandColTot: True 40"
Push A, "GrandRowTot: True"
Push A, "SubTot: AA DD"
Push A, "Wdt: 7: AA BB CC"
Push A, "OutLin: 2: AA BB"
Push A, "OutLin: 3: FF BB"
Push A, "Lbl: AA : AA-Lbl"
Push A, "Lbl: CC : CC-Lbl"
Push A, "Lbl: DD : DD-Lbl"
Push A, "DtaSum: DD Sum #,##0.00"
Push A, "OpnInd: True"
Push A, "Fny: AA BB CC DD EE GG FF"

Dim Fny$()
Dim Act As PtFmtr
Act = PtFmtr(A)
ErBrw Act.Er
AyBrw PtFmtrLy(Act)
Stop
End Sub

Function PtFmtrLy(A As PtFmtr) As String()
Dim O$()
Dim J%
With A
    '**Ori
    PushAy O, ZLy_Ori(.Row, "Row")
    PushAy O, ZLy_Ori(.Pag, "Pag")
    PushAy O, ZLy_Ori(.Dta, "Dta")
    PushAy O, ZLy_Ori(.Col, "Col")
    PushAy O, ZLy_Fmt_Lbl(.LblFld, .LblVal)
    PushAy O, ZLy_Fmt_Wdt(.WdtFld, .WdtVal)
    PushAy O, ZLy_Fmt_OutLin(.OutLinFld, .OutLinLvl)
    PushAy O, ZLy_Tot_DtaSum(.DtaSumFld, .DtaSumFmt, .DtaSumFun)
    PushAy O, ZLy_Tot_SubTot(.SubTotFld)
    PushAy O, ZLy_Tot_GrandColTot(.GrandColTot, .GrandColWdt)
    PushAy O, ZLy_Tot_GrandRowTot(.GrandRowTot)
    PushAy O, ZLy_Tot_OpnInd(.OpnInd)
End With
PtFmtrLy = O
End Function
Function ZLy_Tot_DtaSum(Fld$(), Fmt$(), Fun() As XlConsolidationFunction) As String()

End Function
Function ZLy_Tot_SubTot(Fld$()) As String()
ZLy_Tot_SubTot = ApSy("SubTot : " & Join(Fld, " "))
End Function
Function ZLy_Tot_GrandColTot(Tot As Boolean, Wdt%) As String()
ZLy_Tot_GrandColTot = ApSy("GrandColTot : " & Tot & " " & Wdt)
End Function
Function ZLy_Tot_GrandRowTot(Tot As Boolean) As String()
ZLy_Tot_GrandRowTot = ApSy("GrandRowTot : " & Tot)
End Function
Function ZLy_Tot_OpnInd(OpnInd As Boolean) As String()
ZLy_Tot_OpnInd = ApSy("OpnInd : " & OpnInd)
End Function

Function ZLy_Fmt_Wdt(Fld$(), Wdt%()) As String()

End Function
Function ZLy_Fmt_OutLin(Fld$(), Lvl() As Byte) As String()

End Function
Function ZLy_Fmt_Lbl(Fld$(), Lbl$()) As String()

End Function
Function ZLy_Ori(Fld$(), Ori$) As String()

End Function

Sub PtFmtrTpBrw()
Dim A$()
Push A, "Dim F$()"
Push A, "Push F, ""Fny: """
Push A, "Push F, ""Row: """
Push A, "Push F, ""Col: """
Push A, "Push F, ""Pag: """
Push A, "Push F, ""Dta: """
Push A, "Push F, ""Wdt: 7: """
Push A, "Push F, ""OutLin: 2: """
Push A, "Push F, ""OutLin: 3: """
Push A, "Push F, ""Lbl: AA : """
Push A, "Push F, ""Lbl: CC : """
Push A, "Push F, ""DtaSum: <DtaSumFld> { Sum | Avg | Cnt } <DtaSumFmt>"
Push A, "Push F, ""SubTot: """
Push A, "Push F, ""OpnInd: """
Push A, "Push F, ""GrandColTot: True"""
Push A, "Push F, ""GrandRowTot: True"""
AyBrw A
End Sub

Private Function ZBrk_Fmt_Fmt(S2$, Fny$()) As Variant()
Dim OEr(), OFmtFld$(), OFmtVal$()
Dim F$(), Fmt$
With StrBrk(S2, ":")
    Fmt = .S1
    F = Split(.S2, " ")
End With
Dim I
For Each I In F
    If Not AyHas(Fny, I) Then
        Push OEr, FmtQQ("Fld[?] in Fmt-Lin not found in Fny", I)
    Else
        Push OFmtFld, I
        Push OFmtVal, Fmt
    End If
Next
ZBrk_Fmt_Fmt = Array(OEr, OFmtFld, OFmtVal)
End Function

Private Function ZBrk_Fmt_Lbl(S2$, Fny$(), Dta$()) As Variant()
Dim OEr(), OLblFld$(), OLblDtaFno%(), OLblColFld$(), OLblVal$()
Dim Fld$, Lbl$
With StrBrk(S2, ":")
    Fld = .S1
    Lbl = .S2
End With
Dim Msg$
If Not AyHas(Fny, Fld) Then
    Msg = FmtQQ("Fld [?] with Lbl[?] in Lbl-Lin not found in Fny", Fld, Lbl):
    Push OEr, ErNew(Msg)
Else
    Dim DtaFno%, ColFld$
    If AyHas(Dta, Fld) Then
        DtaFno = AyIdx(Dta, Fld) + 1
    Else
        ColFld = Fld
    End If
    Push OLblFld, Fld
    Push OLblVal, Lbl
    Push OLblDtaFno, DtaFno
    Push OLblColFld, ColFld
End If
ZBrk_Fmt_Lbl = Array(OEr, OLblFld, OLblDtaFno, OLblColFld, OLblVal)
End Function

Private Function ZBrk_Fmt_OutLin(S2$, Fny$()) As Variant()
Dim OEr(), OOutLinFld$(), OOutLinFno%(), OOutLinLvl() As Byte
Dim F$(), Lvl As Byte
With StrBrk(S2, ":")
    Lvl = .S1
    F = Split(.S2, " ")
End With
Dim I
For Each I In F
    If Not AyHas(Fny, I) Then
        Push OEr, FmtQQ("Fld[?] in OutLin-Lin not found in Fny", I)
    Else
        Push OOutLinFld, I
        Push OOutLinLvl, Lvl
        Push OOutLinFno, AyIdx(Fny, I) + 1
    End If
Next
ZBrk_Fmt_OutLin = Array(OEr, OOutLinFld, OOutLinFno, OOutLinLvl)
End Function

Private Function ZBrk_Fmt_Wdt(S2$, Fny$()) As Variant()
Dim OEr(), OWdtFld$(), OWdtFno%(), OWdtVal%()
Dim F$(), Wdt%
With StrBrk(S2, ":")
    Wdt = .S1
    F = Split(.S2, " ")
End With
Dim I
For Each I In F
    If Not AyHas(Fny, I) Then
        Push OEr, FmtQQ("Fld[?] in Wdt not found in Fny", I)
    Else
        Push OWdtFld, I
        Push OWdtFno, AyIdx(Fny, I) + 1
        Push OWdtVal, Wdt
    End If
Next
ZBrk_Fmt_Wdt = Array(OEr, OWdtFld, OWdtFno, OWdtVal)
End Function

Private Function ZBrk_Tot_DtaSum(S2$, Dta$()) As Variant()
Dim OEr(), ODtaSumFld$(), ODtaSumFno%(), ODtaSumFun() As XlConsolidationFunction, ODtaSumFmt$()
'From S2 of fmt :   <DtaSumFld> : { Avg | Cnt | Sum } <DtaSumFmt>
'Where XX is {ODtaSumFld}, and should be found in {Dta} (Dta-Fields)
'{ODtaSumFno} is the Fno (Field-No) of ODtaSumFld in {Dta}
'{ODtaSumFun} is the From { Avg | Cnt | Sum }
Dim SFld$
Dim SFun$
Dim SFmt$
    Dim Ay$()
    Ay = Split(S2, " ")
    If UBound(Ay) <> 2 Then
        Push OEr, "There should be 3 items in in DtaSum-Lin, but now it has [" & UBound(Ay) & "]"
        GoSub OneMoreMsg
        Exit Function
    End If
    SFld = Ay(0)
    SFun = Ay(1)
    SFmt = Ay(2)

Dim Fun As XlConsolidationFunction
    Select Case SFun
    Case "Sum": Fun = XlConsolidationFunction.xlSum
    Case "Cnt": Fun = XlConsolidationFunction.xlCount
    Case "Avg": Fun = XlConsolidationFunction.xlAverage
    Case Else:
        Push OEr, FmtQQ("The <DtaSumFun> [?] element is invalid", SFun)
        GoSub OneMoreMsg
        Exit Function
    End Select
    
If Not AyHas(Dta, SFld) Then
    Push OEr, FmtQQ("<DtaSumFld>[?] in DtaSum-Lin not found in {Dta-Fields}", SFld)
    GoSub OneMoreMsg
    Exit Function
End If
        
Push ODtaSumFld, SFld
Push ODtaSumFun, Fun
Push ODtaSumFno, AyIdx(Dta, SFld) + 1
Push ODtaSumFmt, SFmt
ZBrk_Tot_DtaSum = Array(OEr, ODtaSumFld, ODtaSumFno, ODtaSumFun, ODtaSumFmt)
Exit Function
OneMoreMsg:
    Push OEr, "DtaSum-Lin must in format of [DtaSum : XXX SSS FFF], where XXX is <DtaSumFld>, SSS is {Avg|Sum|Cnt}, FFF is format string.  But now [XXX SSS FFF] is [" & S2 & "]"
End Function

Private Function ZBrk_Tot_GrandColTot(S2$) As Variant()
Dim OEr(), OTot As Boolean, OWdt%
Dim A$(): A = Split(S2, " ")
If Sz(A) <> 2 Then
    Push OEr, "GrandColTot-Lin must have 2 Items: <Bool> <Wdt>, but now it is [" & Sz(A) & "].  S2=[" & S2 & "]"
    Exit Function
End If
On Error GoTo X
OTot = A(0)
On Error GoTo Y
OWdt = A(1)
GoTo Ext
Dim Msg$
X:
    Msg = FmtQQ("GrandColTot-Lin must have 2 Items: <Bool> <Wdt>, now <Bool>[?] cannot convert to boolean", A(0))
    OEr = ErNew(Msg)
    GoTo Ext
Y:
    Msg = FmtQQ("GrandColTot-Lin must have 2 Items: <Bool> <Wdt>, now <Wdt>[?] cannot convert to boolean", A(0))
    OEr = ErNew(Msg)
    GoTo Ext
Ext:
ZBrk_Tot_GrandColTot = Array(OEr, OTot, OWdt)
End Function

Private Function ZBrkBool(S2$, LinPfx$) As Variant()
Dim Er()
On Error GoTo X
ZBrkBool = Array(Er, CBool(S2))
Exit Function
X:
Dim Msg$
Msg = FmtQQ("Lin-[?] must be convertable to boolean", LinPfx$)
ZBrkBool = Array(ErNew(Msg), False)
End Function

Private Function ZFldIdx(Fld$(), Fny$()) As Integer()
If AyIsEmpty(Fld) Then Exit Function
Dim U%
Dim O%()
U = UB(Fld)
ReDim O(U)
Dim J%
For J = 0 To U
    O(J) = AyIdx(Fny, Fld(J)) + 1
Next
ZFldIdx = O
End Function

Private Function ZBrkFld(FnStr$, Fny$(), LinPfx$, IsDup%) As Variant()
Dim OFld$(), OEr()
If IsDup Then
    OEr = ErNew(FmtQQ("Lin-[?] is duplicated", LinPfx))
Else
    OFld = Split(FnStr, " ")
    Dim O$()
    Dim J%
    For J = 0 To UB(OFld)
        If AyHas(Fny, OFld(J)) Then
            Push O, OFld(J)
        Else
            Dim Msg$
            Msg = FmtQQ("Lin-[?] has field [?] not found in Fny", LinPfx, OFld(J))
            PushAy OEr, ErNew(Msg)
        End If
    Next
End If
ZBrkFld = Array(OEr, OFld)
End Function

