Attribute VB_Name = "nAs400_Dtf"
Option Compare Database
Option Explicit

Sub DtfCrt(Dtf$, Sql$, IP$ _
    , Optional Lib$ = "RBPCSF" _
    , Optional IsByXls As Boolean = False _
    , Optional IsRun As Boolean = False _
    , Optional ONrec& = 0 _
    )
'Aim: Build a file [Dtf]  by [{Sql}, {IP}, {Lib}] with optional to run it.
'     which will download data to [FfnDownload] with Fdf in same directory as Dtf.  If no data is download empty Txt or empty Xls will be created according to FDF
'     [mFfnFdf] = Ffnn(Dtf).Fdf
'     [FfnDownload] = Ffnn(Dtf).Txt (or .xls)
'     ResStr @ modResStr.ODtfCxt()
FfnAsstExt Dtf, ".dtf", "DtfCrt"

Dim ODtfCxt$
    ODtfCxt = DtfCxt(Dtf$, IP, Sql, Lib, IsByXls)

FfnDltIfExist Dtf
StrWrt ODtfCxt, Dtf

If IsRun Then
    Dim FfnDownload$
        Dim A$
            A = IIf(IsByXls, ".xls", ".txt")
        FfnDownload = FfnRplExt(Dtf, A)
    DtfRun Dtf, FfnDownload, ONrec
End If
End Sub

Function DtfCrt__Tst()
Dim mNRec&
'If Dtf("Dtf", "AVM_Xls", "Select * from AVM", "192.168.103.14", , , True, True, True, mNRec) Then Stop Else Debug.Print mNRec
'If Dtf("c:\Tmp\IIC.Dtf", "Select * from IIC where iclas='ux'", "192.168.103.13", "BPCSF", False, True, mNRec) Then Stop Else Debug.Print mNRec
DtfCrt "c:\Tmp\IIC.Dtf", "Select * from IIC where iclas='07'", "192.168.103.13", "BPCSF", True, True, mNRec
Debug.Print mNRec
End Function

Function DtfCxt$(Dtf$, IP$, Sql$, Lib$, IsByXls As Boolean)
'Find [ODtfCxt]
'               Txt     Xls
'{IP}
'{LIB}
'{Sql}
'{ConvTy}       0       1
'{Fdf}
'{FfnDownload}
'{PCFilTy}      1       16
'{SavFDF}       1       1

Dim ConvTy%
    ConvTy = IIf(IsByXls, 4, 0)

Dim Fdf$
    Fdf$ = FfnRplExt(Dtf, ".Fdf")
Dim SavFDF$
    SavFDF = "1"
Dim PCFilTy%
    PCFilTy = IIf(IsByXls, 16, 1)
Dim FfnDownload$
    Dim A$
        A = IIf(IsByXls, ".xls", ".txt")
    FfnDownload = FfnRplExt(Dtf, A)

FfnDltIfExist FfnDownload
Dim Tp$
Stop
Tp = ResStr("ODtfCxt", True)
Dim S$
    S = Replace(Replace(Sql, vbLf, " "), vbCr, " ")
    
DtfCxt = FmtNm(Tp, "IP,Lib,Sql,ConvTy,Fdf,FfnDownload,PCFilTy,SavFDF", IP, Lib, S, ConvTy, Fdf, FfnDownload, PCFilTy, SavFDF)
End Function

Function DtfRun(Dtf$, FfnDownload$, Optional ONrec& = 0) As Boolean
'Aim:   Run {Dtf}, which assume to download data to {FfnDownload} & create [mFfnFdf] & return {oNRec}
'Side Effects:
''    Delete:
''      #1 {Dtf}        : will create if error
''    Create & Delete:
''      #2 [mFfnDownload]
''      #3 [mFfnDtfMsg]     : will create if error. (=*.dtf.txt
'Detail:
''       Dlt 2 files: {FfnDownload} & DirOf(FfnDownload]EndDownload
''       Build a Bat file [mFfnBat] to "#1. Run rtopcb with create *.dtf.txt" & "#2 Create EndDownload"
''       Call the bat & wait for [DirOf(FfnDownload]EndDownload] & delete [mFfnBat]
''       Import *.dtf.txt(mFfnDtfMsg) to get {oNRec}
''       create empty {FfnDownload} from [mFfnFdf], if no data is download, and if *.dtf.txt (the dtf download message) said so by return oNRec
''       If not {pKeepDtf}, Rmv {Dtf} & [mFfnDtfMsg]
Const cSub$ = "DtfRun"
'Do build {mFfnBat}, which rtopcs {Dtf}, which assume to download data to {FfnDownload}
Dim mFfnBat$, mFfnDtfMsg$, mFfnDownloadEnd$
Do
    Dim mDir$: mDir = Fct.Nam_DirNam(Dtf)
    mFfnBat = mDir & "Download.Bat"
    mFfnDownloadEnd = mDir & "DownloadEnd"
    If Dlt_Fil(mDir & "EndDownload") Then ss.A 2: GoTo E

    mFfnDtfMsg = Dtf & ".txt"
    Dim mFno As Byte: If Opn_Fil_ForOutput(mFno, mFfnBat, True) Then ss.A 3: GoTo E
    Print #mFno, Fmt_Str("rtopcb /s ""{0}"" >""{1}""", Dtf, mFfnDtfMsg)
    Print #mFno, Fmt_Str("echo >""{0}""", mFfnDownloadEnd)
    Close #mFno
Loop Until True
'Do Run {mFfnBat} to download data to {FfnDownload} & with msg send to {mFfnDtfMsg}
Do
    If Dlt_Fil(FfnDownload) Then ss.A 1: GoTo E
    Shell """" & mFfnBat & """", vbHide
    If Fct.WaitFor(mFfnDownloadEnd, "[" & FfnDownload & "] <--Downloading File" & vbCrLf & "[" & Dtf & "] <--By Dtf File") Then ss.A 1, "User has cancelled to wait": GoTo E
    Dlt_Fil mFfnBat
Loop Until True

'Find {oNRec} from {mFfnDtfMsg}
Stop
'If Run_RecCnt_ByFfnDtfMsg(oNRec, mFfnDtfMsg) Then ss.A 4: GoTo E

'Do create empty {FfnDownload} from mFfnFdf, if no data is download, and if *.dtf.txt (the dtf download message) said so by return oNRec,
Dim mFfnFdf$: mFfnFdf = Cut_Ext(Dtf) & ".Fdf"
Do
    If VBA.Dir(FfnDownload) = "" Then
        If ONrec > 0 Then ss.A 5, "No FfnDownload is found, but NRec>0", eImpossibleReachHere: GoTo E
        Select Case Right(FfnDownload, 4)
        Case ".xls"
            Dim mWb As Workbook: If Crt_Xls_FmFDF(FfnDownload, mFfnFdf) Then ss.A 6: GoTo E
        Case ".txt"
            If Opn_Fil_ForOutput(mFno, FfnDownload) Then ss.A 7: GoTo E
            Close #mFno
        Case Else
            ss.A 8, "No FfnDownload is found and it is not .Xls or .Txt": GoTo E
        End Select
    End If
Loop Until True
If ONrec > 0 Then Dlt_Fil Dtf
Exit Function
R: ss.R
E:
End Function

Sub DtfRun__Tst()
Const cFfnDtf$ = "c:\tmp\aa.dtf"
Dim mNRec&: DtfCrt cFfnDtf, "Select * from IIC where iclas='x12'", "192.168.103.14", , True, True, mNRec
End Sub

