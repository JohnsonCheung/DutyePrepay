Attribute VB_Name = "nXls_nObj_nPt_nInf_Pt"
Option Compare Database
Option Explicit

Function PtNmNz$(PtNm$, A As Worksheet)
If PtNm <> "" Then PtNmNz = PtNm: Exit Function
Dim J%
For J = 1 To 1000
    If Not PtNmIsExist("PT_" & J, A) Then PtNmNz = "PT_" & J: Exit Function
Next
Er "PtNmNz: Impossible"
End Function

Function PtToStr$(A As PivotTable)
If IsNothing(A) Then PtToStr = "#Nothing#": Exit Function
Dim mCmdTxt$: mCmdTxt = "CmdTxt<Nil>"
Dim mCnnStr$: mCnnStr = "CnnStr<Nil>"
Dim mPcRfhNm$: mPcRfhNm = "PcRfhNm<Nil>"
Dim mPcIdx%
On Error Resume Next
With A
    mCmdTxt = .PivotCache.CtCommandText
    mCnnStr = .PivotCache.Connection
    mPcRfhNm = .PivotCache.RefreshName
    mPcIdx = .PivotCache.Index
End With
On Error GoTo 0
PtToStr = LpApToStr(CtComma, "CmdTxt,PcIdx,PtNam,PcRfhNm,CnnStr", mCmdTxt, mPcIdx, A.Name, mPcRfhNm, mCnnStr)
End Function
