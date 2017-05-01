Attribute VB_Name = "mCmd_FrmKE24CmdImp"
Option Compare Database
Option Explicit

Sub CmdKE24Clear(pY As Byte, pM As Byte)
If VdtYr(pY) Then Exit Sub
If VdtMth(pM) Then Exit Sub
If Not Start("Clear sales history data (KE24) for Year[" & pY + 2000 & "] Month[" & pY & "]?", "Clear?") Then Exit Sub
Dim mCndn$: mCndn = Fmt("Yr={0} and Mth={1}", pY, pM)
SqlRun "Delete From KE24 where " & mCndn
SqlRun "Update KE24H set NCopaOrd=0,NCopaLin=0,NCus=0,NSKU=0,Qty=0,Tot=0,DteUpd=Now() where " & mCndn
Done
End Sub

Sub CmdKE24Clear__Tst()
CmdKE24Clear 9, 2
End Sub

Sub CmdKE24Import(pY As Byte, pM As Byte)
If VdtYr(pY) Then Exit Sub
If VdtMth(pM) Then Exit Sub
Dim mLsFx$: mLsFx = GetLsFx_KE24(pY, pM): If mLsFx = "" Then Exit Sub
If Not Start("Following files are found, Import?" & vbLf & vbLf & mLsFx, "Import?") Then Exit Sub
Dim mAyFx$(): mAyFx = Split(mLsFx, vbLf)
Dim J%
DoCmd.SetWarnings False
For J = 0 To UBound(mAyFx)
    CmdKE24Import_1ImpFx mAyFx(J), pY, pM
Next
CmdKE24Import_2UpdKE24H pY, pM
Done
End Sub

Sub CmdKE24Import__Tst()
CmdKE24Import 10, 1
End Sub

Private Sub CmdKE24Import_1ImpFx(Fx$, pY As Byte, pM As Byte)
'Aim: Import Fx$ into KE24 of pY, pM
If Dir(Fx) = "" Then MsgBox "make sure this exists" & vbLf & Fx: Exit Sub
If TblCrt_FmLnkWs(Fx, "Sheet1", TNew:=">KE24") Then MsgBox "Cannot create link table [>KE24] to Xls" & vbLf & Fx: Exit Sub
With CurrentDb.OpenRecordset("Select Count(*) from [>KE24] where Year([Posting Date])<>" & 2000 + pY & " or Month([Posting Date])<>" & pM)
    If Nz(.Fields(0).Value, 0) <> 0 Then
        MsgBox "There are [" & .Fields(0).Value & "] records with Posting Date not in " & pY + 2000 & "/" & pM & vbLf & vbLf & Fx, vbCritical, "Error in import file"
        .Close
        Exit Sub
    End If
End With
CmdKE24Import_1ImpKE24         ' Import >KE24 to KE24
CmdKE24Import_2UpdKE24H pY, pM ' Update current pY, pM of KE24H by KE24
End Sub

Private Sub CmdKE24Import_1ImpKE24()
SqlRun "SELECT [Document number]                     AS CopaNo," & _
" CLng(IIf(Trim(Nz([Item number],''))='',0,[Item number])) AS CopaLNo," & _
                                          " [Posting date] AS PostDate," & _
                                 " CStr(Nz([Product],'-')) AS Sku," & _
                                  " CStr(Nz([Customer],0)) AS Cus," & _
                       " CLng(Nz(-[Billing qty in SKU],0)) AS Qty," & _
                              " CCur(Nz([D&T invoiced],0)) AS Tot" & _
" INTO [#KE24] FROM [>KE24];"
SqlRun "INSERT INTO KE24 (CopaNo,   CopaLNo,   PostDate,   Sku,   Cus,   Qty,   Tot, Yr,                    Mth)" & _
                     " SELECT x.CopaNo, x.CopaLNo, x.PostDate, x.Sku, x.Cus, x.Qty, x.Tot, Year(x.PostDate)-2000, Month(x.PostDate)" & _
                     " FROM [#KE24] x " & _
                     " LEFT JOIN KE24 a ON x.CopaNo=a.CopaNo AND x.CopaLNo=a.CopaLNo" & _
                     " WHERE a.CopaNo Is Null;"
End Sub

Private Sub CmdKE24Import_2UpdKE24H(pY As Byte, pM As Byte)
SqlRun Fmt("SELECT Yr,Mth,Sum(x.Qty) AS Qty, Sum(x.Tot) AS Tot, Count(CopaLNo) AS NLin INTO [#Sum]        FROM KE24 x        WHERE Yr={0} And Mth={1} GROUP BY Yr,Mth;", pY, pM)
SqlRun Fmt("SELECT Yr,Mth, CopaNo                                                      INTO [#SumOrdList] FROM KE24          WHERE Yr={0} And Mth={1} GROUP BY Yr,Mth,CopaNo;", pY, pM)
        SqlRun "SELECT Yr,Mth, Count(CopaNo) AS NOrd                                       INTO [#SumOrdCnt]  FROM [#SumOrdList]                          GROUP BY Yr,Mth;"
SqlRun Fmt("SELECT Yr,Mth, Cus                                                         INTO [#SumCusList] FROM KE24          WHERE Yr={0} And Mth={1} GROUP BY Yr,Mth,Cus;", pY, pM)
SqlRun Fmt("SELECT Yr,Mth, Count(Cus) AS NCus                                          INTO [#SumCusCnt]  FROM [#SumCusList]                          GROUP BY Yr,Mth;", pY, pM)
SqlRun Fmt("SELECT Yr,Mth, Sku                                                         INTO [#SumSkuList] FROM KE24          WHERE Yr={0} And Mth={1} GROUP BY Yr,Mth,Sku;", pY, pM)
SqlRun Fmt("SELECT Yr,Mth, Count(Sku) AS NSku                                          INTO [#SumSkuCnt]  FROM [#SumSkuList]                          GROUP BY Yr,Mth;", pY, pM)
SqlRun Fmt("UPDATE (((KE24H x" & _
                              " INNER JOIN [#Sum]       a ON x.Mth=a.Mth AND x.Yr=a.Yr)" & _
                              " INNER JOIN [#SumCusCnt] b ON x.Mth=b.Mth AND x.Yr=b.Yr)" & _
                              " INNER JOIN [#SumOrdCnt] c ON x.Mth=c.Mth AND x.Yr=c.Yr)" & _
                              " INNER JOIN [#SumSkuCnt] d ON x.Mth=d.Mth AND x.Yr=d.Yr" & _
                              " SET x.Qty     =a.Qty, x.Tot=a.Tot, x.NCopaLin=a.NLin," & _
                                  " x.NCus    =b.NCus," & _
                                  " x.NCopaOrd=c.NOrd," & _
                                  " x.NSku    =d.NSku, x.DteUpd = Now()" & _
                              " WHERE x.Qty     <>a.Qty Or x.Tot<>a.Tot Or x.NCopaLin<>a.NLin" & _
                                 " Or x.NCus    <>b.NCus" & _
                                 " Or x.NCopaOrd<>c.NOrd" & _
                                 " Or x.NSku    <>d.NSku" _
                                 , pY, pM)
End Sub

Private Function GetLsFx_KE24$(pY As Byte, pM As Byte)
'Aim: Get list of Fx separated by VbLf in .\Import\KE24 yyyy-mm*.xls
Dim mDir$: mDir = FbCurPth & "SAPDownloadExcel\"
Dim mFxSpec$: mFxSpec = mDir & "KE24 " & pY + 2000 & "-" & Format(pM, "00") & "*.xls"
Dim mA$: mA = Dir(mFxSpec): If mA = "" Then MsgBox "No such file found:" & vbLf & vbLf & mFxSpec: Exit Function
mA = mDir & mA
Dim mB$: mB = Dir
While mB <> ""
    mA = mA & vbLf & mDir & mB
    mB = Dir
Wend
GetLsFx_KE24 = mA
End Function

