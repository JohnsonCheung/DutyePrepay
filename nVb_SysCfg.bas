Attribute VB_Name = "nVb_SysCfg"
Option Compare Text
Option Explicit
'------------------------
'-- Fm SdirWrkObj & "Cfg.Txt
Private cfgDirTmp$
Private cfgDirExp$
Private cfgDirImp$
Private cfgApp$
Private cfgDsn$
Private cfgIsLclMd As Boolean
Private cfgIsNoLogin As Boolean
Private cfgIsDbg As Boolean
Private cfgIsDbgOdbc As Boolean
Private cfgLgcHidAcs As Boolean
Private cfgIsDbgRunAcs As Boolean
Private cfgOdbcTimeOut%

Function SysCfg_DirExp$()
SysCfg_zReadCfg
SysCfg_DirExp = cfgDirExp
End Function

Function SysCfg_DirImp$()
SysCfg_zReadCfg:                 SysCfg_DirImp = cfgDirImp:
End Function

Function SysCfg_IsDbg() As Boolean
SysCfg_zReadCfg: SysCfg_IsDbg = cfgIsDbg
End Function

Function SysCfg_IsDbgOdbc() As Boolean:
SysCfg_zReadCfg:
SysCfg_IsDbgOdbc = cfgIsDbgOdbc:
End Function

Function SysCfg_IsDbgRunAcs() As Boolean
SysCfg_zReadCfg
SysCfg_IsDbgRunAcs = cfgIsDbgRunAcs:
End Function

Function SysCfg_IsLclMd() As Boolean:
SysCfg_zReadCfg:
SysCfg_IsLclMd = cfgIsLclMd
End Function

Function SysCfg_IsNoLogin() As Boolean:
SysCfg_zReadCfg: SysCfg_IsNoLogin = cfgIsNoLogin

End Function

Function SysCfg_LgcHidAcs() As Boolean:
SysCfg_zReadCfg
SysCfg_LgcHidAcs = cfgLgcHidAcs
End Function

Function SysCfg_OdbcTimeOut%():
SysCfg_zReadCfg:
SysCfg_OdbcTimeOut = cfgOdbcTimeOut%:
End Function

Private Function SysCfg_zReadCfg() As Boolean
Const cSub$ = "zReadCfg"
Static xIsReadCfg As Boolean
If xIsReadCfg Then Exit Function
xIsReadCfg = True
Dim mFfn$: mFfn = Sdir_PgmObj & "Cfg.txt"
If VBA.Dir(mFfn, vbHidden) = "" Then MsgBox "No cfg.txt in PgmObj dir": Application.Quit
Dim mFno As Byte: If Opn_Fil_ForInput(mFno, mFfn) Then Application.Quit
cfgDirTmp = "c:\Tmp\"
cfgDirExp = "c:\Tmp\Export\"
cfgDirExp = "c:\Tmp\Import\"
cfgIsDbgRunAcs = True

While Not EOF(mFno)
    Dim mL$: Line Input #mFno, mL
    If Left(mL, 1) = "#" Then GoTo Nxt
    Dim mK$, mV$
    With Brk(mL, "=")
        mK = .S1
        mV = .S2
    End With
    Select Case mK
    Case "DirTmp": cfgDirTmp = mV
    Case "App": cfgApp = mV
    Case "Dsn": cfgDsn = mV
    Case "OdbcTimeOut": cfgOdbcTimeOut = mV
    Case "IsLclMd": cfgIsLclMd = mV
    Case "IsNoLogin": cfgIsNoLogin = mV
    Case "IsDbg": cfgIsDbg = mV
    Case "IsDbgOdbc": cfgIsDbgOdbc = mV
    Case "DirExp": cfgDirExp = mV
    Case "DirImp": cfgDirImp = mV
    Case "cfgLgcHidAcs": cfgLgcHidAcs = mV
    Case "cfgIsDbgRunAcs": cfgIsDbgRunAcs = mV
    End Select
Nxt:
Wend
Close #mFno
Crt_Dir cfgDirTmp
Crt_Dir cfgDirExp
Crt_Dir cfgDirImp
Exit Function
R: ss.R
E:
End Function
