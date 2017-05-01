Attribute VB_Name = "nAcs_nCtl_ToStr"
Option Compare Database
Option Explicit

Function CtlsToStr$(Ctls As Access.Controls, Optional pWithTag As Boolean, Optional pSepChr$ = CtComma)
On Error GoTo R
Dim mS$, iCtl As Access.Control
For Each iCtl In Ctls
    mS = Push(mS, CtlToStr(iCtl, pWithTag), pSepChr)
Next
CtlsToStr = mS
Exit Function
R: CtlsToStr = "Err: CtlToStrs(Ctls).  Msg=" & Err.Description
End Function

Function CtlToStr$(Ctl As Access.Control, Optional WithTag As Boolean)
On Error GoTo R
If WithTag Then
    If IsNothing(Ctl.Tag) Then
        CtlToStr = Ctl.Name
    Else
        CtlToStr = Ctl.Name & "(" & Ctl.Tag & ")"
    End If
Else
    CtlToStr = Ctl.Name
End If
Exit Function
R: CtlToStr = "Err: CtlToStr(Ctl).  Msg=" & Err.Description
End Function
