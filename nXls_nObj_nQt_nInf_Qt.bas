Attribute VB_Name = "nXls_nObj_nQt_nInf_Qt"
Option Compare Database
Option Explicit

Function QtToStr$(A As QueryTable)
If IsNothing(A) Then QtToStr = "#Nothing#": Exit Function
Dim mCmdTxt$: mCmdTxt = "CmdTxt<Nil>"
Dim mCnnStr$: mCnnStr = "CnnStr<Nil>"
On Error Resume Next
With A
    mCmdTxt = .CtCommandText
    mCnnStr = .Connection
End With
On Error GoTo 0
QtToStr = LpApToStr(CtComma, "CmdTxt,QtNam,CnnStr", mCmdTxt, A.Name, mCnnStr)
End Function
