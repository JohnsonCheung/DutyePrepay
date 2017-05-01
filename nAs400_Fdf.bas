Attribute VB_Name = "nAs400_Fdf"
Option Compare Database
Option Explicit

Function FdfCrtFx(FxTar$, pFfnFdf$) As Boolean
Const cSub$ = "FdfCrtFx"
'Aim: Create a 2 rows Xls {FxTar} from {pFfnFDF}.  All numeric fields will set zero.
'PCFDF
'PCFT 16
'PCFO 1, 1, 5, 1, 1
'PCFL IID 20 4
'PCFL ICLAS 20 4
'PCFL ICDES 20 60
'PCFL ICGL 20 40
'PCFL ICCOGA 20 40
'PCFL ICTAX 20 10
'PCFL ICPPGL 20 40
'PCFL ICALGL 20 40
'PCFL ICALC1 20 20
'PCFL ICALC2 20 20
'PCFL ICALC3 20 20
'PCFL ICALC4 20 20
'PCFL ICALC5 20 20
'PCFL ICALG1 20 40
'PCFL ICALG2 20 40
'PCFL ICALG3 20 40
'PCFL ICALG4 20 40
'PCFL ICALG5 20 40
'PCFL ICALP1 2 7/2
'PCFL ICALP2 2 7/2
'PCFL ICALP3 2 7/2
'PCFL ICALP4 2 7/2
'PCFL ICALP5 2 7/2
'PCFL ICRETA 20 40
'PCFL ICCORA 20 40
'PCFL ICLMPC 2 7/2
'PCFL ICUMPC 2 7/2
On Error GoTo R
Dim mF As Byte: If Opn_Fil_ForInput(mF, pFfnFdf) Then ss.A 1: GoTo E
Dim mL$, mA
Line Input #mF, mL: mA = "PCFDF":          If mL <> mA Then ss.A 2, "[" & mA & "] is expected", , "Current Line Value", mL: GoTo E
Line Input #mF, mL: mA = "PCFT 16":        If mL <> mA Then ss.A 3, "[" & mA & "] is expected", , "Current Line Value", mL: GoTo E
Line Input #mF, mL: mA = "PCFO 1,1,5,1,1": If mL <> mA Then ss.A 4, "[" & mA & "] is expected", , "Current Line Value", mL: GoTo E
Dim mWb As Workbook: If Crt_Wb(mWb, FxTar) Then ss.A 5, "Cannot create FxTar": GoTo E
Dim mDir$, mFnn$, mExt$: If Brk_Ffn_To3Seg(mDir, mFnn, mExt, FxTar) Then ss.A 6: GoTo E
mWb.Sheets(1).Name = mFnn
Dim J%
For J = mWb.Worksheets.Count - 1 To 2 Step -1
    mWb.Worksheets(J).Delete
Next
J = 1
Dim mWs As Worksheet: Set mWs = mWb.Worksheets(1)
While Not EOF(mF)
    Line Input #mF, mL: If Left(mL, 5) <> "PCFL " Then ss.A 7, "[PCFL ] is expected", , "Current Line Value", mL: GoTo E
    Dim mX$(): mX = Split(mL)
    mWs.Cells(1, J).Value = mX(1)
    Select Case mX(2)
    Case "2": mWs.Cells(2, J).Value = 0
    End Select
    J = J + 1
Wend
Cls_Wb mWb, True
Close #mF
Exit Function
R: ss.R
E:
End Function

Function FdfCrtFx__Tst()
Const cFfnFdf$ = "c:\aa.fdf"
Const cFx = "c:\aa.xls"
Dim mFno As Byte: If Opn_Fil_ForOutput(mFno, cFfnFdf, True) Then Stop
Print #mFno, "PCFDF"
Print #mFno, "PCFT 16"
Print #mFno, "PCFO 1,1,5,1,1"
Print #mFno, "PCFL IID 20 4"
Print #mFno, "PCFL ICLAS 20 4"
Print #mFno, "PCFL ICDES 20 60"
Print #mFno, "PCFL ICGL 20 40"
Print #mFno, "PCFL ICCOGA 20 40"
Print #mFno, "PCFL ICTAX 20 10"
Print #mFno, "PCFL ICPPGL 20 40"
Print #mFno, "PCFL ICALGL 20 40"
Print #mFno, "PCFL ICALC1 20 20"
Print #mFno, "PCFL ICALC2 20 20"
Print #mFno, "PCFL ICALC3 20 20"
Print #mFno, "PCFL ICALC4 20 20"
Print #mFno, "PCFL ICALC5 20 20"
Print #mFno, "PCFL ICALG1 20 40"
Print #mFno, "PCFL ICALG2 20 40"
Print #mFno, "PCFL ICALG3 20 40"
Print #mFno, "PCFL ICALG4 20 40"
Print #mFno, "PCFL ICALG5 20 40"
Print #mFno, "PCFL ICALP1 2 7/2"
Print #mFno, "PCFL ICALP2 2 7/2"
Print #mFno, "PCFL ICALP3 2 7/2"
Print #mFno, "PCFL ICALP4 2 7/2"
Print #mFno, "PCFL ICALP5 2 7/2"
Print #mFno, "PCFL ICRETA 20 40"
Print #mFno, "PCFL ICCORA 20 40"
Print #mFno, "PCFL ICLMPC 2 7/2"
Print #mFno, "PCFL ICUMPC 2 7/2"
Close #mFno
If Ovr_Wrt(cFx, True) Then Stop
If FdfCrtFx(cFx, cFfnFdf) Then Stop
Dim mWb As Workbook: If Opn_Wb_RW(mWb, cFx, , True) Then Stop
End Function

Function FdfFny(Fdf$) As String()
'Aim: Cv {pFfnFdf} to oLnFld.  If will be either * or [list of field] as yymd_ prefix field converted
'FDF forAt:
'PCFDF
'PCFT 1
'PCFO 1,1,5,1,1
'PCFL IID 1 2
'PCFL IPROD 1 15
Dim F%: F = FtOpnInp(Fdf)
Dim L$
Line Input #F, L: If L <> "PCFDF" Then Er "{Line-1} of {Fdf} must be [PCFDF]", L, Fdf
Line Input #F, L: If L <> "PCFT 16" And L <> "PCFT 1" Then Er "{Line 2} of {Fdf} must be [PCFT 16] or [PCFT 1]", L, Fdf
Line Input #F, L: If L <> "PCFO 1,1,5,1,1" Then Er "{Line-3} of {Fdf} must be [PCFO 1,1,5,1,1]", L, Fdf
Dim O$(), Lno%
Lno = 2
While Not EOF(F)
    Lno = Lno + 1
    Line Input #F, L: If Left(L, 5) <> "PCFL" Then Er "{Line-4} {Lno} onward of {Fdf} must begin with [PCFL]", L, Lno, Fdf
    Dim Ay$(): Ay = Split(L)
    Push O, Ay(1)
Wend
Close #F
FdfFny = O
End Function

Function FdfFny__Tst()
Const cFfnDtf$ = "c:\aa.dtf"
Const cFfnFdf$ = "c:\aa.fdf"
DtfCrt cFfnDtf, "Select IIC.*,ICUMPC AS yymd_ICUMPC  from IIC", "192.168.103.14", , , True
AyDmp FdfFny(cFfnFdf)
End Function

Sub FdfWrtSchemaIni(Fdf$)
'Aim: Create [ScheA.ini] in the same directory as {Fdf}
'FDF forAt:
'PCFDF
'PCFT 1
'PCFO 1,1,1,1,1
'PCFL IID 1 2
'PCFL IPROD 1 15
'ScheA.ini ForAt:
'[IIM_Short.txt]
'ColNameHeader = False
'ForAt = FixedLength
'AxScanRows = 100
'CharacterSet = OEM
'Col1="IID" Char Width 2
'Col2="IPROD" Char Width 15
Dim Fnn$
    Fnn = FfnFnn(Fdf)
     'mDir$, mExt$: If jj.Brk_Ffn_To3Seg(mDir, Fnn, mExt, Fdf) Then ss.A 1: GoTo E

Dim F%: F = FtOpnInp(Fdf)
    Dim L$, A$
    A = "PCFDF":           Line Input #F, L: If L <> A Then Er "{Line-1}must be {This}", L, A
    A = "PCFT 1":          Line Input #F, L: If L <> A Then Er "{Line-2}must be {This}", L, A
    A = "PCFO 1,1,5,1,1":  Line Input #F, L: If L <> A Then Er "{Line-3}must be {This}", L, A

Dim O$()
    Push O, "[" & Fnn & ".txt]"
    Push O, "ColNameHeader = False"
    Push O, "ForAt = FixedLength"
    Push O, "AxScanRows = 100"
    Push O, "CharacterSet = OEM"
    Dim N%
    While Not EOF(F)
        N = N + 1
        Line Input #F, L
        Dim Ay$()
            Ay = Split(L)
        
        If Ay(0) <> "PCFL" Then Er "Line#" & N & " does not begin with PCFL", L
        
        Dim Nm$
        Dim W%
            Nm = Ay(1): W = Val(Ay(3))
        Dim T$
            Select Case Ay(2)
            Case "1": T = "Char"
            Case "2"
                If InStr(Ay(3), "/") > 0 Then
                    T = "Double"
                Else
                    Select Case W
                    Case Is <= 2: T = "Byte"
                    Case Is <= 4: T = "Integer"
                    Case Is <= 9: T = "Long"
                    Case Else
                        T = "Double"
                    End Select
                End If
            End Select
        Push O, FmtQQ("Col?=""?"" ? Width ?", N, Nm, T, W)
    Wend
Dim OFt$
    OFt = FfnPth(Fdf) & "ScheA.ini"

AyWrt O, OFt
End Sub

Function FdfWrtSchemaIni__Tst()
FdfWrtSchemaIni "D:\Data\Johnson Cheung\MyDoc\My Projects\My Projects Library\Ldb\Ldb\WorkingDir\Data\IIM.fdf"
End Function
