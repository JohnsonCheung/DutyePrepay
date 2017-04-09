Attribute VB_Name = "mCmd_FrmYrOCmdBldOpnBal"
Option Compare Database
Option Explicit
Option Base 0

Sub FrmYrOCmdBldOpnBal(Y As Byte)
'Aim: Build Year Opening (table YrOD) by Y
'     Case#1: If not Last Year data, try import
'     Case#2: If there is last year data, build the YrOD and Update YrO
If VdtYr(Y) Then Exit Sub
If Not IsLasYrOD_Exist(Y) Then
    If Not Start("This is the first year opening." & vbLf & vbLf & "Import from Excel?") Then Exit Sub
    CmdBldOpn_1ImpFirstYrOpn Y
    Exit Sub
End If
If Not Start("Start building Year Opening [" & Y + 2000 & "]") Then Exit Sub
SqlRun FmtQQ("SELECT ? AS Yr, Sku, OpnQty AS Q, OpnTot AS A INTO [#Mge]             FROM YrOD WHERE Yr=?", Y, Y - 1)
SqlRun FmtQQ("INSERT INTO [#Mge] (Yr,Sku,Q,A) SELECT ?, Sku, Sum(-Qty) , Sum(-Tot)  FROM KE24 WHERE Yr=? GROUP BY Sku;", Y, Y - 1)
SqlRun FmtQQ("INSERT INTO [#Mge] (Yr,Sku,Q,A) SELECT ?, Sku, Sum(x.Qty), Sum(x.Amt) FROM PermitD x INNER JOIN Permit ON x.Permit = a.Permit WHERE Year(PostDate)-2000)=? GROUP BY Sku;", Y, Y - 1)
SqlRun FmtQQ("INSERT INTO [#Mge] (Yr,Sku,A)   SELECT ?, Sku, AdjTot                 FROM YrAdjD WHERE Yr=?;", Y, Y - 1)

SqlRun FmtQQ("DELETE FROM YrOD WHERE Yr=?", Y)
        SqlRun "INSERT INTO YrOD (Yr,Sku,OpnQty,OpnTot) SELECT Yr, Sku, Sum(Q), Sum(A) FROM [#Mge] GROUP BY Yr,Sku;"
SqlRun FmtQQ("UPDATE YrOD SET OpnRate = OpnTot/OpnQty WHERE OpnQty<>0 AND Yr=?", Y)
'------
SqlRun FmtQQ("SELECT Yr, Count(1) AS NSku, Sum(x.OpnQty) AS OpnQty, Sum(x.OpnTot) AS OpnTot INTO [#Tot] FROM YrOD WHERE Yr=? GROUP BY Yr;", Y)
SqlRun "UPDATE YrO x INNER JOIN [#Tot] a ON a.Yr = x.Yr SET x.NSku=a.NSku, x.OpnQty = a.OpnQty, x.OpnTot = a.OpnTot, x.DteUpd = Now();"


'CmdBldOpn_2YrOD Y  ' Build current Y YrOD from last year YrOD+In-Out+Adj
'CmdBldOpn_3YrO Y   ' Update YrO by YrOD of current Y
End Sub

Private Sub CmdBldOpn_1ImpFirstYrOpn(Y As Byte)
Dim mYear%: mYear = Y + 2000
Dim mFfn$: mFfn = ImpFdr & "Duty Prepay Year Opening " & mYear & ".xls"
If Dir(mFfn) = "" Then MsgBox "make sure this exists" & vbLf & mFfn: Exit Sub
Stop
'If TblCrt_FmLnkWs(mFfn, "Sheet1", pNmtNew:=">FirstYrOpn") Then MsgBox "Cannot create link table [>FirstYrOpn]": Exit Sub
'Aim: Import >FirstYrOpn into YrOD
        SqlRun "SELECT Trim(CStr(x.SKU)) AS SKU, Sum(x.Qty) AS Qty, Sum(x.Amt) as Amt INTO [#Inp] FROM [>FirstYrOpn] x GROUP BY Trim(CStr(SKU))"
        SqlRun "DELETE FROM YrOD WHERE Yr=" & Y
SqlRun FmtQQ("INSERT INTO YrOD (Sku,OpnQty,OpnTot,Yr) SELECT SKU, Sum(Qty), Sum(Amt), ? FROM [#Inp] GROUP BY SKU;", Y)
SqlRun FmtQQ("UPDATE YrOD SET OpnRate =OpnTot/OpnQty WHERE Yr=?;", Y)
SqlRun FmtQQ("SELECT Yr, Count(1) AS NSku, Sum(x.OpnQty) AS OpnQty, Sum(x.OpnTot) AS OpnTot INTO [#Tot] FROM YrOD x WHERE Yr=? GROUP BY Yr;", Y)
SqlRun "UPDATE YrO x INNER JOIN [#Tot] a ON a.Yr=x.Yr SET x.NSku=a.NSku, x.OpnQty=a.OpnQty, x.OpnTot=a.OpnTot, x.DteUpd= Now();"
End Sub

Private Function IsLasYrOD_Exist(Y As Byte) As Boolean
Dim mYr As Byte
mYr = Y - 1
With CurrentDb.OpenRecordset("Select count(*) from YrOD where Yr=" & mYr)
    If Nz(.Fields(0).Value, 0) > 0 Then IsLasYrOD_Exist = True
    .Close
End With
End Function
