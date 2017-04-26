Attribute VB_Name = "nXls_nFmt_PtFmtr"
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
End Type

Function PtFmtr(PtFmtrLy$()) As PtFmtr
Dim L, A As S1S2, S2$, Err$()
Dim O As PtFmtr
'-------------
Dim Fny$()
    For Each L In PtFmtrLy
        If Trim(L) = "" Then GoTo Nxt1
        A = StrBrk(L, ":")
        S2 = A.S2
        Select Case A.S1
        Case "Fny": Fny = Split(S2, " "): Exit For
        End Select
Nxt1:
    Next
    If Sz(Fny) = 0 Then Er "no {Fny} is found in {PtFmtrLy}"

    With O
        .Fny = Fny
        For Each L In PtFmtrLy
            If Trim(L) = "" Then GoTo Nxt
            A = StrBrk(L, ":")
            S2 = A.S2
            Select Case A.S1
            Case "Fny":
            Case "Lbl": ZBrk_Fmt_Lbl S2, Err, Fny, .Dta, .LblFld, .LblDtaFno, .LblColFld, .LblVal
            Case "Row": .Row = ZBrkFld(S2, Err, Fny, "Row", Sz(.Row) > 0)
            Case "Col": .Col = ZBrkFld(S2, Err, Fny, "Col", Sz(.Col) > 0)
            Case "Pag": .Pag = ZBrkFld(S2, Err, Fny, "Pag", Sz(.Pag) > 0)
            Case "Dta": .Dta = ZBrkFld(S2, Err, Fny, "Dta", Sz(.Dta) > 0)
            Case "Wdt":    ZBrk_Fmt_Wdt S2, Err, Fny, .WdtFld, .WdtFno, .WdtVal
            Case "OutLin": ZBrk_Fmt_OutLin S2, Err, Fny, .OutLinFld, .OutLinFno, .OutLinLvl
            Case "DtaSum": ZBrk_Tot_DtaSum S2, Err, .Dta, .DtaSumFld, .DtaSumFno, .DtaSumFun, .DtaSumFmt
            Case "SubTot": .SubTotFld = ZBrkFld(S2, Err, Fny, "SubTot", Sz(.SubTotFld) > 0)
            Case "OpnInd": .OpnInd = ZBrkBool(S2, Err, "OpnInd")
            Case "GrandColTot": ZBrk_Tot_GrandColTot S2, Err, .GrandColTot, .GrandColWdt
            Case "GrandRowTot": .GrandRowTot = ZBrkBool(S2, Err, "GrandRowTot")
            Case Else: Push Err, "Lin [" & L & "] has invalid type.  Valid Type are [Lbl Row Col Pag Dta Fmt Wdt OutLin SubTot DtaSum GrandColTot GrandRowTot]"
            End Select
Nxt:
        Next
    End With
With O
    .SubTotFno = ZFldIdx(.SubTotFld, Fny)
End With
PtFmtr = O
If Not AyIsEmpty(Err) Then
    Push Err, "Fny: " & Join(Fny, " ")
    Push Err, "DtaFld: " & Join(O.Dta, " ")
    AyBrw Err
End If
End Function

Sub PtFmtr__Tst()
Dim A$()
Push A, "Row: AA BB CC"
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
AyBrw PtFmtrLy(Act)
Stop
End Sub

Function PtFmtrLy(A As PtFmtr) As String()
Dim O$()
With A
    Push O, "Row: " & Join(.Row, " ")
    Push O, "Col: " & Join(.Col, " ")
    Push O, "Pag: " & Join(.Pag, " ")
    Push O, "Dta: " & Join(.Dta, " ")
    Dim J%
    For J = 0 To UB(.LblFld)
        Push O, "Lbl: " & .LblFld(J) & ": " & .LblVal(J)
    Next
    For J = 0 To UB(.OpnInd)
        Push O, "Lbl: " & .LblFld(J) & ": " & .LblVal(J)
    Next
    End With
PtFmtrLy = O
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

Sub ZBrk_Fmt_Fmt(S2$, OEr$(), Fny, OFmtFld$(), OFmtVal$())
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
End Sub

Sub ZBrk_Fmt_Lbl(S2$, OEr$(), Fny$(), Dta$(), OLblFld$(), OLblDtaFno%(), OLblColFld$(), OLblVal$())
Dim Fld$, Lbl$
With StrBrk(S2, ":")
    Fld = .S1
    Lbl = .S2
End With
If Not AyHas(Fny, Fld) Then Push OEr, FmtQQ("Fld [?] with Lbl[?] in Lbl-Lin not found in Fny", Fld, Lbl): Exit Sub
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
End Sub

Sub ZBrk_Fmt_OutLin(S2$, OEr$(), Fny$(), OOutLinFld$(), OOutLinFno%(), OOutLinLvl() As Byte)
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
End Sub

Sub ZBrk_Fmt_Wdt(S2$, OEr$(), Fny, OWdtFld$(), OWdtFno%(), OWdtVal%())
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
End Sub

Sub ZBrk_Tot_DtaSum(S2$, OEr$(), Dta$(), ODtaSumFld$(), ODtaSumFno%(), ODtaSumFun() As XlConsolidationFunction, ODtaSumFmt$())
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
        Exit Sub
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
        Exit Sub
    End Select
    
If Not AyHas(Dta, SFld) Then
    Push OEr, FmtQQ("<DtaSumFld>[?] in DtaSum-Lin not found in {Dta-Fields}", SFld)
    GoSub OneMoreMsg
    Exit Sub
End If
        
Push ODtaSumFld, SFld
Push ODtaSumFun, Fun
Push ODtaSumFno, AyIdx(Dta, SFld) + 1
Push ODtaSumFmt, SFmt
Exit Sub
OneMoreMsg:
    Push OEr, "DtaSum-Lin must in format of [DtaSum : XXX SSS FFF], where XXX is <DtaSumFld>, SSS is {Avg|Sum|Cnt}, FFF is format string.  But now [XXX SSS FFF] is [" & S2 & "]"
End Sub

Sub ZBrk_Tot_GrandColTot(S2$, OEr$(), OTot As Boolean, OWdt%)
Dim A$(): A = Split(S2, " ")
If Sz(A) <> 2 Then
    Push OEr, "GrandColTot-Lin must have 2 Items: <Bool> <Wdt>, but now it is [" & Sz(A) & "].  S2=[" & S2 & "]"
    Exit Sub
End If
On Error GoTo X
OTot = A(0)
On Error GoTo Y
OWdt = A(1)
Exit Sub
X:
    Push OEr, FmtQQ("GrandColTot-Lin must have 2 Items: <Bool> <Wdt>, now <Bool>[?] cannot convert to boolean", A(0))
    Exit Sub
Y:
    Push OEr, FmtQQ("GrandColTot-Lin must have 2 Items: <Bool> <Wdt>, now <Wdt>[?] cannot convert to boolean", A(0))
    Exit Sub
End Sub

Function ZBrkBool(S2$, OEr$(), Msg$) As Boolean
On Error GoTo X
ZBrkBool = S2
Exit Function
X: Push OEr, "Line [" & Msg & "] must be convertable to boolean"
End Function

Function ZFldIdx(Fld$(), Fny$()) As Integer()
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

Private Function ZBrkFld(FnStr$, OEr$(), Fny$(), Msg$, IsDup%) As String()
If IsDup Then
    Push OEr, "Lin [" & Msg & "] is duplicated: " & FnStr
    Exit Function
End If
Dim F$(): F = Split(FnStr, " ")
Dim O$()
Dim J%
For J = 0 To UB(F)
    If AyHas(Fny, F(J)) Then
        Push O, F(J)
    Else
        Push OEr, Msg & ": has field [" & F(J) & "] not found in Fny"
    End If
Next
ZBrkFld = O
End Function
