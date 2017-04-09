Attribute VB_Name = "nFs_Ffn"
Option Compare Database
Option Explicit

Function FfnAddFnSfx$(Ffn, Sfx)
FfnAddFnSfx = FfnCutExt(Ffn) & Sfx & FfnExt(Ffn)
End Function

Sub FfnAsstExist(Ffn, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst FfnChkExist(Ffn), Av
End Sub

Sub FfnAsstExt(Ffn, Ext, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst FfnChkExt(Ffn, Ext), Av
End Sub

Sub FfnAsstNotExist(Ffn, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst FfnChkNotExist(Ffn), Av
End Sub

Function FfnChkExist(Ffn) As Dt
If Not FfnIsExist(Ffn) Then FfnChkExist = ErNew("{Ffn} unexpectedly not exist", Ffn)
End Function

Function FfnChkExt(Ffn, Ext) As Dt
If FfnExt(Ffn) <> Ext Then FfnChkExt = ErNew("{Ffn} should have {Ext}", Ffn, Ext)
End Function

Function FfnChkNotExist(Ffn) As Dt
If FfnIsExist(Ffn) Then FfnChkNotExist = ErNew("{Ffn} unexpectedly exists", Ffn)
End Function

Sub FfnCpy(Fm, ToFfn, Optional OvrWrt As Boolean)
FfnOvrWrt ToFfn, OvrWrt
Fso.CopyFile Fm, ToFfn
End Sub

Function FfnCutExt$(Ffn)
Dim P%: P = InStrRev(Ffn, ".")
If P = 0 Then FfnCutExt = Ffn: Exit Function
FfnCutExt = Left(Ffn, P - 1)
End Function

Sub FfnDlt(Ffn)
On Error GoTo R
Kill Ffn
Exit Sub
R: Er "Cannot Delete {file} with {Reason}", Ffn, Err.Description
End Sub

Sub FfnDltIfExist(Ffn)
If FfnIsExist(Ffn) Then FfnDlt Ffn
End Sub

Function FfnEnsPth$(Fil)
Dim O$
If FfnHasPth(Fil) Then
    O = Fil
Else
    O = CurDir & "\" & Fil
End If
FfnEnsPth = O
End Function

Function FfnExt$(Ffn)
With StrBrk1FmEnd(Ffn, ".")
    If InStr(.S2, "\") Then Exit Function
    FfnExt = "." & .S2
End With
End Function

Function FfnFn$(Ffn)
Dim P%: P = InStrRev(Ffn, "\"): If P = 0 Then FfnFn = Ffn: Exit Function
FfnFn = Mid(Ffn, P + 1)
End Function

Function FfnFnn$(Ffn)
FfnFnn = FfnCutExt(FfnFn(Ffn))
End Function

Function FfnHasPth(Fil) As Boolean
FfnHasPth = InStr(Fil, "\") > 0
End Function

Function FfnIsExist(Ffn) As Boolean
FfnIsExist = Fso.FileExists(Ffn)
End Function

Function FfnIsToDay(Ffn$) As Boolean
FfnIsToDay = (Date = CDate(Format(VBA.FileDateTime(Ffn), "yyyy/mm/dd")))
End Function

Sub FfnMov(Ffn, ToPth$)
Fso.MoveFile Ffn, ToPth
End Sub

Sub FfnOvrWrt(Ffn, Optional OvrWrt As Boolean)
'If Ffn not exist, just return
'If OvrWrt, kill it
'Er telling Ffn exist
Dim Exist As Boolean
    Exist = FfnIsExist(Ffn)
If Not Exist Then Exit Sub
If OvrWrt Then
    FfnDlt Ffn
Else
    Er "FfnOvrWrt: {File} exist", Ffn
End If
End Sub

Function FfnPth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\"): If P = 0 Then Exit Function
FfnPth = Left(Ffn, P)
End Function

Sub FfnRenBackup(Ffn, BackupLvl%)
Stop
End Sub

Function FfnRplExt$(Ffn, Ext$)
FfnRplExt = FfnCutExt(Ffn) & Ext
End Function

Function FfnWaitFor(Ffn, Optional Msg$ = "") As Boolean
'Aim: for a file is created.  Return true if "wait for" success.  If cancel waiting by user return false.
Const cSub$ = "WaitFor"
Dim A As FrmOpt: A = FrmOpn("frmWaitFor", ApJnComma(1000, Ffn, Msg, True))
If Not A.Som Then Exit Function
If Not FfnIsExist(Ffn) Then Er "Impossible: FrmOpn('frmWaitFor') returns no error, but Ffn not found."
FfnDltIfExist Ffn
FfnWaitFor = True
End Function
