Attribute VB_Name = "nDao_SchemaIni"
Option Compare Database
Option Explicit
Enum eSchemaIniFmt
    eCsv = 1
End Enum
Enum eSchemaIniFldTy
    eBit = 1
    eByt = 2
    eChr = 3
    eCur = 4
    eDte = 5
    eFlt = 6
    eInt = 7
    eLCh = 8
    eLng = 9
    eSht = 10
    eSng = 11
End Enum

Type TSchemaIniFld
    Nm As String
    Ty As eSchemaIniFldTy
    Wdt As Integer
End Type

Type TSchemaIni
    FcsvFn As String
    Fmt As eSchemaIniFmt
    MaxScanRows As Integer
    Col() As TSchemaIniFld
End Type

Function SchemaIniFldNew(Nm$, Ty As eSchemaIniFldTy, Optional Wdt%) As TSchemaIniFld
Dim O As TSchemaIniFld
With O
    .Nm = Nm
    .Ty = Ty
    .Wdt = Wdt
End With
SchemaIniFldNew = O
End Function

Sub SchemaIniWrt(FschemaIni$, A() As TSchemaIni)
Dim J%
Dim O$()
For J = 0 To UBound(A)
    PushAy O, ZOupLy(A(J))
Next
AyWrt O, FschemaIni
End Sub

Function ZSampleTSchemaIni() As TSchemaIni
Dim O(1) As TSchemaIni
Dim M As TSchemaIni
Dim Col() As TSchemaIniFld


'O(0) ----------------------------------
    ReDim Col(4)
    Col(0) = SchemaIniFldNew("AA", eChr, 10)
    Col(1) = SchemaIniFldNew("BB", eInt)
    Col(2) = SchemaIniFldNew("CC", eDte)
    Col(3) = SchemaIniFldNew("DD", eBit)
    Col(4) = SchemaIniFldNew("EE", eLng)

    M.Col = Col
    M.FcsvFn = "ABC.csv"
    M.Fmt = eCsv
    M.MaxScanRows = 25
    O(0) = M
'O(0) ----------------------------------
    ReDim Col(4)
    Col(0) = SchemaIniFldNew("AA", eChr, 10)
    Col(1) = SchemaIniFldNew("BB", eInt)
    Col(2) = SchemaIniFldNew("CC", eDte)
    Col(3) = SchemaIniFldNew("DD", eBit)
    Col(4) = SchemaIniFldNew("EE", eLng)

    M.Col = Col
    M.FcsvFn = "ABC.csv"
    M.Fmt = eCsv
    M.MaxScanRows = 25
    O(0) = M
'O(0) ----------------------------------
    ReDim Col(4)
    Col(0) = SchemaIniFldNew("AA", eChr, 10)
    Col(1) = SchemaIniFldNew("BB", eInt)
    Col(2) = SchemaIniFldNew("CC", eDte)
    Col(3) = SchemaIniFldNew("DD", eBit)
    Col(4) = SchemaIniFldNew("EE", eLng)

    M.Col = Col
    M.FcsvFn = "ABC.csv"
    M.Fmt = eCsv
    M.MaxScanRows = 25
    O(0) = M


'[A.csv]
'ColNameHeader = True
'Format = CSVDelimited
'MaxScanRows = 25
'CharacterSet = ANSI
'Col1=AA Char Width 255
'Col2=BB Integer
'Col3=CC Integer
'Col4=DD Date
'[f.csv]
'ColNameHeader = True
'Format = CSVDelimited
'MaxScanRows = 25
'CharacterSet = ANSI
'Col1=AA Bit
'Col2=BB Byte
'Col3=CC Char Width 1
'Col4=DD Currency
'Col5=E Date
'Col6=F Float
'Col7=G Integer
'Col8=H LongChar
'Col9=I Short
'Col10=J Single

End Function

Private Function ZFldTyToStr$(A As eSchemaIniFldTy)
Dim O$
Select Case A
Case eSchemaIniFldTy.eBit: O = "Bit"
Case eSchemaIniFldTy.eByt: O = "Byte"
Case eSchemaIniFldTy.eChr: O = "Char"
Case eSchemaIniFldTy.eCur: O = "Currency"
Case eSchemaIniFldTy.eDte: O = "Date"
Case eSchemaIniFldTy.eFlt: O = "Float"
Case eSchemaIniFldTy.eInt: O = "Integer"
Case eSchemaIniFldTy.eLCh: O = "LongChar"
Case eSchemaIniFldTy.eSht: O = "Short"
Case eSchemaIniFldTy.eSng: O = "Single"
Case Else: Stop
End Select
ZFldTyToStr = O
End Function

Private Function ZFmtToStr$(A As eSchemaIniFmt)
Dim O$
Select Case A
Case eSchemaIniFmt.eCsv
    O = "CSVDelimited"
Case Else
    Stop
End Select
ZFmtToStr = O
End Function

Private Function ZOupFld$(A As TSchemaIniFld, Base1Idx%)
Const C$ = "Col? = ??"
Dim W$: W = IIf(A.Ty = eSchemaIniFldTy.eChr Or A.Ty = eSchemaIniFldTy.eLCh, " Width " & A.Wdt, "")
Dim T$: T = ZFldTyToStr(A.Ty)
ZOupFld = FmtQQ(C, A.Nm, T, W)
End Function

Private Function ZOupLy(A As TSchemaIni) As String()
Dim O$()
Push O, FmtQQ("[?]", A.FcsvFn)
Push O, "Format = " & ZFmtToStr(A.Fmt)
Dim J%
For J = 0 To UBound(A.Col)
    Push O, ZOupFld(A.Col(J), J + 1)
Next
ZOupLy = O
End Function
