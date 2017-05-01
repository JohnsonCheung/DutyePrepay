Attribute VB_Name = "nXls_nObj_nPc_Pc"
Option Compare Database
Option Explicit

Function PcToStr$(A As PivotCache)
If IsNothing(A) Then PcToStr = "#Nothing#": Exit Function
On Error GoTo R
Dim mCmdTxt$: mCmdTxt = "CmdTxt<Nil>"
Dim mCnnStr$: mCnnStr = "CnnStr<Nil>"
Dim mRfhNam$: mRfhNam = "RfhNam<Nil>"
Dim mPcIdx%
On Error Resume Next
With A
    mCmdTxt = .CtCommandText
    mCnnStr = .Connection
    mRfhNam = .RefreshName
    mPcIdx = .Index
End With
PcToStr = LpApToStr(CtComma, "CmdTxt,PcIdx,RfhNam,CnnStr", mCmdTxt, mPcIdx, mRfhNam, mCnnStr)
Exit Function
R: PcToStr = ErStr("PcToStr")
End Function
