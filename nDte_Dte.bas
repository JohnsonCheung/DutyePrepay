Attribute VB_Name = "nDte_Dte"
Option Compare Database
Option Explicit

Function Dte2Fy$(Optional pDte As Date = 0)
Dte2Fy = FyNoToStr(Dte2FyNo(pDte))
End Function

Function Dte2FyNo(Optional pDte As Date = 0) As Byte
Dim mDte As Date: mDte = IIf(pDte = 0, Date, pDte)
If Month(mDte) = 1 Then Dte2FyNo = Year(mDte) - 2000: Exit Function
Dte2FyNo = Year(mDte) - 1999
End Function

Function Dte2Qtr(Optional pDte As Date = 0) As String
Dim mDte As Date: mDte = IIf(pDte = 0, Date, pDte)
Dim mMM%: mMM = Month(mDte)
If mMM = 1 Then
    Dte2Qtr = "Q4"
Else
    Dte2Qtr = "Q" & Int((mMM - 2) / 3) + 1
End If
End Function

Function DteAsk( _
      Optional Dft As Date = 0 _
    , Optional Min As Date = #1/1/1980# _
    , Optional Max As Date = #12/31/2100# _
    , Optional IsAlwTim As Boolean _
    , Optional IsAlwNull As Boolean _
    ) As OptDte
Const cSub$ = "Dte"
Dim A As Date: A = IIf(Dft = 0, Date, Dft)
Dim Opt As FrmOpt: Opt = FrmOpnOpt("frmSelDte", ApJnComma(CtComma, Dft, Min, Max, IsAlwTim, IsAlwNull), True)
DteAsk = Form_frmSelDte.RetOptDte
End Function

Function DteAsk___Tst()
Const cSub$ = "Dte_Tst"
Dim mDte As Date, mIsNull As Boolean
Dim mDteDef As Date: mDteDef = CDate(InputBox("Default Date", , Date))
Dim A As OptDte: A = DteAsk(mDteDef, , , True, True)
If Not A.Som Then MsgBox "Select date is cancelled": Exit Function
MsgBox LpApToStr(vbLf, "Selected Date, IsNull", mDte, mIsNull)
End Function

Function DteLasWkLasDte(A As Date) As Date
Dim mWeekday As Byte: mWeekday = Weekday(A, vbSunday) ' Sunday count as first day of a week & week day of Sunday is 1 & Saturday (last date of a week) is 7
DteLasWkLasDte = A - mWeekday
End Function

Function DteWkNo(A As Date) As Byte
If Year(A) = 2005 Then
    DteWkNo = VBA.Format(A, "ww", , vbFirstFullWeek)
    Exit Function
End If
DteWkNo = VBA.Format(A, "ww")
End Function

Function YrWkToStr$(Yr As Byte, Wk As Byte)
YrWkToStr = "Yr" & Format(Yr, "00") & "Wk" & Format(Wk, "00")
End Function

