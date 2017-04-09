Attribute VB_Name = "nXls_Appx"
Option Compare Database
Option Explicit
Private X_DspAlert() As Boolean
Private X_Xls() As Excel.Application
Private X_Appx As Excel.Application

Function Appx() As Excel.Application
On Error GoTo X
Dim Nm$: Nm = X_Appx.Name
Set Appx = X_Appx
Exit Function
X: Set X_Appx = New Excel.Application
Set Appx = X_Appx
End Function

Sub AppxArgeH(Optional A As Excel.Application)
AppxNz(A).Windows.Arrange xlArrangeStyleHorizontal
End Sub

Sub AppxArgeV(Optional A As Excel.Application)
AppxNz(A).Windows.Arrange xlArrangeStyleVertical
End Sub

Sub AppxClsWbs(Optional A As Excel.Application, Optional NoSav As Boolean)
Dim I As Workbook
For Each I In AppxNz(A).Workbooks
    WbCls I, NoSav
Next
End Sub

Sub AppxMinAllWin(Optional A As Excel.Application)
Dim I As Window
For Each I In AppxNz(A).Windows
    If I.WindowState <> xlMinimized Then I.WindowState = xlMinimized
Next
End Sub

Function AppxNz(A As Excel.Application)
If IsNothing(A) Then
    Set AppxNz = Appx
Else
    Set AppxNz = A
End If
End Function

Sub AppxQuit()
On Error Resume Next
X_Appx.DisplayAlerts = False
X_Appx.Quit
Set X_Appx = Nothing
End Sub

Sub AppxSavWb()
Dim iWb As Workbook, iWin As Window
For Each iWb In Excel.Application.Workbooks
    If Not iWb.Saved Then
        iWb.Windows(1).WindowState = xlMaximized
        iWb.Save
    End If
    Set_Wb_Min iWb
Next
End Sub

Sub AppxSetAllWbMin()
Dim iWb As Workbook, iWin As Window
For Each iWb In Excel.Application.Workbooks
    Set_Wb_Min iWb
Next
End Sub

Function CrtPjx(PjNm$, Optional Pth$) As vbproject
Dim F$:  F = PjNm_NewFalm(PjNm, Pth)
Set CrtPjx = FxlamCrt(F)
End Function

Sub CrtPjx__Tst()
Dim P$: P = PjPth
Dim N$: N = TmpFil
Dim F$: F = PjNm_NewFalm(N)
FfnDltIfExist F
CrtPjx N, TmpPth
End Sub

Sub FxlamCrt__Tst()
End Sub

Function NzAppx(Xls As Excel.Application) As Excel.Application
If IsNothing(Xls) Then
    Set NzAppx = Excel.Application
Else
    Set NzAppx = Xls
End If
End Function

Function OpnPjFxlam(Fxlam$) As vbproject
FfnAsstExt Fxlam, ".xlam", "OpnPjx"
Appx.Workbooks.Open Fxlam
Set OpnPjFxlam = VbeLasPj(Appx.Vbe)
End Function

Sub XlsDspAlertPop()
Dim X As Excel.Application
Set X = PopObj(X_Xls)
X.DisplayAlerts = Pop(X_DspAlert)
End Sub

Sub XlsDspAlertPush(Optional Xls As Excel.Application, Optional DspAlert As Boolean)
PushObj X_Xls, NzAppx(Xls)
Push X_DspAlert, Xls.DisplayAlerts
Xls.DisplayAlerts = DspAlert
End Sub
