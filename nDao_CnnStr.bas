Attribute VB_Name = "nDao_CnnStr"
Option Compare Database
Option Explicit
Type Cnn
    TblNm As String
    CnnStr As String
    AppNm As String
    Ver As Byte
    Ext As String
    Msg As String
End Type

Function CnnBrk(CnnStr$) As Cnn
'Brk   ";DATABASE=N:\SapAccessReports\DutyPrepay5\DutyPrepay5_Data.accdb"
'Into AppNm Ver Ext CnnStr Msg
'Skip TblNm
Const CC1$ = "SAPAccessReports\"
Const CC2$ = "Universe\"
Dim O As Cnn:   O.CnnStr = CnnStr
Dim OMsg$, OVer As Byte, oExt$, OAppNm$
If InStr(CnnStr, "SAPAccessReports\") > 0 Then:  CnnBrk_WithSAPAccessReports CnnStr, OAppNm, OVer, oExt: GoTo X1
If InStr(CnnStr, "Universe\") > 0 Then:          CnnBrk_WithUniverse CnnStr, OAppNm, OVer, oExt: GoTo X1
OMsg = FmtQQ("CnnStr.Brk({CnnStr}) must contain [SAPAccessReports\] or [Universe\]", CnnStr): GoTo X2
X1:
O.AppNm = OAppNm
O.Ver = OVer
O.Ext = oExt
X2:
O.Msg = OMsg
CnnBrk = O
End Function

Function CnnStr_Csv$(pFfnCsv)
'Text;DSN=Delta_Tbl_08052203_20080522_033948 Link Specification;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE=C:\Tmp;TABLE=Delta_Tbl_08052203_20080522_033948#csv
End Function

Function CnnStr_Xls$(Pfx$)
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
CnnStr_Xls = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & Pfx & ";"
End Function

Function CnnStrDr(A$) As Variant()
Dim O(5)
Dim B$
Dim M As Cnn
With Brk(A, "|")
    O(0) = .S1
    M = CnnBrk(.S2)
    O(1) = M.AppNm
    O(2) = M.Ver
    O(3) = M.Ext
    O(4) = M.Msg
    O(5) = M.CnnStr
End With
CnnStrDr = O
End Function

Function CnnStrFx$(LnkFxCnnStr$)
Dim Cnn$: Cnn = LnkFxCnnStr
If Left(Cnn, 10) <> "Excel 8.0;" Then Exit Function
Dim P%: P = InStr(Cnn, "DATABASE="): If P <= 0 Then Exit Function
CnnStrFx = Mid(Cnn, P + 9)
End Function

Function CnnStrFx__Tst()
Const cFfn$ = "c:\Book1.xls"
Dim mWb As Workbook: If Crt_Wb(mWb, cFfn, True) Then Stop
Cls_Wb mWb, True
If TblCrt_FmLnkXls(cFfn) Then Stop
Dim mFfn$, mA$
mA = "Sheet1": If Fnd_Ffn_Fm_LnkXlsNmt(mFfn, mA) Then Stop Else Debug.Print mA, mFfn
mA = "Sheet2": If Fnd_Ffn_Fm_LnkXlsNmt(mFfn, mA) Then Stop Else Debug.Print mA, mFfn
mA = "Sheet3": If Fnd_Ffn_Fm_LnkXlsNmt(mFfn, mA) Then Stop Else Debug.Print mA, mFfn
End Function

Function CnnStrLnkFb$(Fb$)
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
CnnStrLnkFb = ";DATABASE=" & Fb
End Function

Function CnnStrLnkFbOle$(Fb$)
'    "Provider=Microsoft.JET.OLEDB.4.0;"
CnnStrLnkFbOle = Fmt( _
    "OLEDB;" & _
    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "User ID=Admin;" & _
    "Data Source={0};" & _
    "Mode=Share Deny None;" & _
    "Jet OLEDB:Engine Type=5;" & _
    "Jet OLEDB:Database Locking Mode=1;" & _
    "Jet OLEDB:Global Partial Bulk Ops=2;" & _
    "Jet OLEDB:Global Bulk Transactions=1;" & _
    "Jet OLEDB:Create System Database=False;" & _
    "Jet OLEDB:Encrypt Database=False;" & _
    "Jet OLEDB:Don't Copy Locale on Compact=False;" & _
    "Jet OLEDB:Compact Without Replica Repair=False;" & _
    "Jet OLEDB:SFP=False", Fb)
End Function

Function CnnStrLnkFx$(Fx$)
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
CnnStrLnkFx = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & Fx & ";"
End Function

Private Sub CnnBrk__Tst()
Dim CnnStr$:
    CnnStr = CurrentDb.TableDefs("Permit").Connect
Dim Act As Cnn
    Act = CnnBrk(CnnStr)
With Act
    Debug.Assert .AppNm = "DutyPrepay"
    Debug.Assert .CnnStr = ";DATABASE=N:\SapAccessReports\DutyPrepay5\DutyPrepay5_Data.accdb"
    Debug.Assert .Ext = ".accdb"
    Debug.Assert .Msg = ""
    Debug.Assert .TblNm = ""
    Debug.Assert .Ver = "5"
End With
End Sub

Private Sub CnnBrk_WithSAPAccessReports(CnnStr$, OAppNm$, OVer As Byte, oExt$)
Dim A$: A = TakAft(CnnStr, "SAPAccessReports\")
Dim AppSeg$, Fn$
With Brk(A, "\")
    AppSeg = .S1
    Fn = .S2
End With
OVer = 0
Dim D$: D = Right(AppSeg, 1)
If ChrIsDig(D) Then
    OVer = D
    OAppNm = RmvLasChr(AppSeg)
Else
    OAppNm = AppSeg
End If
oExt = FfnExt(Fn)
End Sub

Private Sub CnnBrk_WithUniverse(CnnStr$, OAppNm$, OVer As Byte, oExt$)
Dim A$: A = TakAft(CnnStr, "Universe\")
OAppNm = "Unverise"
OVer = 0
oExt = FfnExt(A)
End Sub

