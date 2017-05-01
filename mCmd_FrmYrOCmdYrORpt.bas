Attribute VB_Name = "mCmd_FrmYrOCmdYrORpt"
Option Compare Database
Option Explicit

Sub FrmYrOCmdYrORpt(Y As Byte)
'Aim Gen an Excel report with Month and Year by Y & pM
DoCmd.SetWarnings False
If Y > Year(Date) Then MsgBox "Y[" & Y & "] cannot > current year[" & Year(Date) & "].", vbCritical: Exit Sub
Dim mFxTo$: mFxTo = zzFxYrRpt(Y)
Dim mWb As Workbook
If Dir(mFxTo) <> "" Then
    If Not Start("Report exist, Regenerate?") Then
        Set mWb = FxWb(mFxTo)
        mWb.Application.Visible = True
        Exit Sub
    End If
End If

' Create @RptY & @RptM
'Aim: Create table @RptY @RptM from Permit,PermitD,KE24
'YpIO NmYpIO
'1   Opn
'2   In
'3   Out
'4   Close
'5   Adjusted
'6   New Clos
TmpMge Y  ' Create #TmpMge
OupRptM       ' Create @RptM from #TmpMge
OupRptY Y    ' Create @RptY

Dim mFxFm$: mFxFm = zzFxYrRptTp
FfnCpy mFxFm, mFxTo
Dim mWs As Worksheet
Set mWb = FxWb(mFxTo)       ' The Tp contain query to @QryRptY & @QryRptM which will put additional columns to table @RptY & @RptM
WbRfh mWb
WbSav mWb
WbVis mWb
End Sub

Sub RunCmdYrORpt__Tst()
FrmYrOCmdYrORpt 10
End Sub

Private Sub Oup__Tst()
DoCmd.SetWarnings False
TmpMge 10
OupRptM
OupRptY 10
End Sub

Private Sub OupRptM()
SqlRun "Delete from `@RptM`"
SqlRun "Insert into `@RptM` Select * from `#TmpMge`"
End Sub

Private Sub OupRptY(Y As Byte)
OupRptY_1TmpMgeIO
OupRptY_2MaxClsMth Y
SqlRun "DELETE FROM [@RptY] WHERE Yr=" & Y
SqlRun "INSERT INTO [@RptY] (Yr,Sku,YpIO,Q,A) SELECT   Yr,  Sku,  YpIO,  Q,  A FROM [@RptM]  WHERE YpIO=1 AND Mth=1 and Yr=" & Y
SqlRun "INSERT INTO [@RptY] (Yr,Sku,YpIO,Q,A) SELECT   Yr,  Sku,  YpIO,  Q,  A FROM [#TmpMgeIO] WHERE Yr=" & Y
SqlRun "INSERT INTO [@RptY] (Yr,Sku,YpIO,Q,A) SELECT x.Yr,x.Sku,x.YpIO,x.Q,x.A FROM [@RptM] x INNER JOIN [#MaxClsMth] a ON x.Yr=a.Yr AND x.Mth=a.MaxClsMth AND x.YpIO=a.YpIO WHERE x.Yr=" & Y
SqlRun "INSERT INTO [@RptY] (Yr,Sku,YpIO,Q,A) SELECT   Yr,  Sku,  YpIO,  Q,  A FROM [@RptM]  WHERE YpIO In (5,6) and Yr=" & Y
End Sub

Private Sub OupRptY_1TmpMgeIO()
SqlRun "SELECT Yr,Sku,YpIO,Sum(x.Q) AS Q, Sum(x.A) AS A INTO [#TmpMgeIO] FROM [#TmpMge] x WHERE YpIO In (2,3) GROUP BY Yr,Sku,YpIO;"
End Sub

Private Sub OupRptY_2MaxClsMth(Y As Byte)
SqlRun FmtQQ("SELECT Yr,YpIO,Max(Mth) AS MaxClsMth INTO [#MaxClsMth] FROM [@RptM] WHERE Yr=? And YpIO=4 GROUP BY Yr,YpIO;", Y)
End Sub

Private Sub TmpMge(Y As Byte)
TmpMge_1YrOD Y       ' Create #TmpMge from YrOD
TmpMge_2In Y         ' Insert #TmpMge from PermitD as In
TmpMge_3Out Y        ' Insert #TmpMge from KE24    as Out
TmpMge_4JanCls        ' Insert #TmpMge from Jan:Opn/In/Out as Jan:Cls
TmpMge_5OpnCls Y     ' Insert #TmpMge from Feb-Dec:Opn/Cls
TmpMge_6AdjYrD Y     ' Insert #TmpMge from AdjYrD
TmpMge_7NewCls        ' Insert #TmpMge from #TmpMge for NewCls
SqlRun "Delete from [#TmpMge] where Nz(A,0)=0 and Nz(Q,0)=0"
SqlRun "UPDATE [#TmpMge] SET A=Null WHERE A=0;"
SqlRun "UPDATE [#TmpMge] SET Q=Null WHERE Q=0;"
TmpMge_8AllYpIO Y
End Sub

Private Sub TmpMge_1YrOD(Y As Byte)
SqlRun FmtQQ("SELECT Yr, CByte(1) AS Mth, Sku, CByte(1) AS YpIO, Sum(OpnQty) AS Q, Sum(OpnTot) AS A INTO [#TmpMge] FROM YrOD Where Yr=? GROUP BY Yr,Sku", Y)
End Sub

Private Sub TmpMge_2In(Y As Byte)
SqlRun FmtQQ("INSERT INTO `#TmpMge` (Yr,Mth,Sku,YpIO,Q,A)" & _
" SELECT Year(PostDate)-2000, Month(PostDate), Sku, 2, Sum(a.Qty), Sum(a.Amt)" & _
" FROM Permit x INNER JOIN PermitD a ON a.Permit = x.Permit" & _
" Where Year(PostDate)-2000=?" & _
" GROUP BY Year(PostDate)-2000, Month(PostDate), Sku", Y)
End Sub

Private Sub TmpMge_3Out(Y As Byte)
SqlRun Fmt("INSERT INTO [#TmpMge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,Mth,Sku,3,Sum(Qty),Sum(Tot) FROM KE24 x WHERE Yr={0} GROUP BY Yr,Mth,Sku", Y)
End Sub

Private Sub TmpMge_4JanCls()
SqlRun "INSERT INTO [#TmpMge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,1,Sku,4,Sum(Q),Sum(A) FROM [#TmpMge] WHERE Mth=1 GROUP BY Yr,Sku;"
End Sub

Private Sub TmpMge_5OpnCls(Y As Byte)
Dim J%
For J = 2 To zM(Y) ' If Y is current year, return current month else return 12
    SqlRun Fmt("INSERT INTO [#TmpMge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,{0},Sku,1,Sum(Q),Sum(A) FROM [#TmpMge] WHERE Mth={1} and YpIO=4 GROUP BY Yr,Sku;", J, J - 1)
    SqlRun Fmt("INSERT INTO [#TmpMge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,{0},Sku,4,Sum(Q),Sum(A) FROM [#TmpMge] WHERE Mth={0}            GROUP BY Yr,Sku;", J)
Next
End Sub

Private Sub TmpMge_6AdjYrD(Y As Byte)
'Aim: Insert #TmpMge from AdjYrD
SqlRun Fmt("INSERT INTO [#TmpMge] (Yr,Mth,Sku,YpIO,A) SELECT Yr,13,Sku,5,Sum(AdjTot) FROM YrAdjD WHERE Yr={0} GROUP BY Yr,Sku;", Y)
End Sub

Private Sub TmpMge_7NewCls()
SqlRun "INSERT INTO [#TmpMge] (Yr,Mth,Sku,YpIO,Q,A) SELECT Yr,12,Sku,6,Sum(Q),Sum(A) FROM [#TmpMge] WHERE YpIO In (4,5) AND Mth=12 GROUP BY Yr,Sku;"
End Sub

Private Sub TmpMge_8AllYpIO(Y As Byte)
'Aim: Write record to #TmpMge for ( 1st SKU x 12 month x 6 YpIO )
Dim mSku$: mSku = TmpMge_8AllYpIO_1MinSku
Dim J%, I%
With CurrentDb.TableDefs("#TmpMge").OpenRecordset
    For J = 1 To zM(Y)
        For I = 1 To 6
            .AddNew
            !Yr = Y
            !Mth = J
            !Sku = mSku
            !YpIO = I
            !Q = 0
            !A = 0
            .Update
        Next
    Next
    .Close
End With
End Sub

Private Function TmpMge_8AllYpIO_1MinSku$()
With CurrentDb.OpenRecordset("Select Min(Sku) from `#TmpMge`")
    TmpMge_8AllYpIO_1MinSku = .Fields(0).Value
    .Close
End With
End Function

Private Function zM(Y As Byte) As Byte
If Y + 2000 = Year(Date) Then
    zM = Month(Date)
Else
    zM = 12
End If
End Function

Private Function zzFxYrRpt$(Y As Byte)
zzFxYrRpt = FbCurPth & "Output\Duty prepay report - Year " & Y + 2000 & ".xls"
End Function

Private Function zzFxYrRptTp$()
zzFxYrRptTp = FbCurPth & "Template\Template_DutyPrepay_Report.xls"
End Function

