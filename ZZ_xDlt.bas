Attribute VB_Name = "ZZ_xDlt"
Option Compare Text
Option Explicit
Option Base 0

Function Dlt_Fil_ByAy(pDir$, pAyFn$()) As Boolean
Const cSub$ = "Dlt_Fil_ByAy"
Dim J%
For J = 0 To Sz(pAyFn) - 1
    If Dlt_Fil(pDir & pAyFn(J)) Then ss.A 1: GoTo E
Next
Exit Function
E:
End Function

Function Dlt_Fil_ByPfx(pDir$, pPfx$) As Boolean
Const cSub$ = "Dlt_Fil_ByPfx"
Dim mAyFn$()
If Fnd_AyFn_ByLik(mAyFn, pDir, pPfx & "*") Then ss.A 1: GoTo E
If Dlt_Fil_ByAy(pDir, mAyFn) Then ss.A 2: GoTo E
Exit Function
E:
End Function

Function Dlt_Rel(pNmRel$, Optional pDb As database) As Boolean
Const cSub$ = "Dlt_Rel"
On Error GoTo R
DbNz(pDb).Relations.Delete pNmRel
R: ss.R
E:
End Function

Function Dlt_RelAll(Optional pDb As database) As Boolean
Dim mDb As database: Set mDb = DbNz(pDb)
With mDb.Relations
    While .Count >= 1
        .Delete mDb.Relations(0).Name
    Wend
End With
End Function

Function Dlt_RelAll__Tst()
Const cFbMeta$ = "C:\Tmp\WorkingDir\Meta_Data.Mdb"
Dim mDb As database
If Opn_Db(mDb, cFbMeta, False) Then Stop
If Dlt_RelAll(mDb) Then Stop
mDb.Close
If Opn_CurDb(G.gAcs, cFbMeta) Then Stop
G.gAcs.Visible = True
End Function

Function Dlt_RowNotInAy(Rg As Range, pAy$()) As Boolean
'Aim: for all data downward from {Rg} delete any row having value not in {pAy}
Const cSub$ = "Dlt_RowNotInAy"
On Error GoTo R
Dim mRnoLas&: If Fnd_RnoLas(mRnoLas, Rg) Then ss.A 1: GoTo E
Dim iRCnt&
For iRCnt = mRnoLas - Rg.Row + 1 To 1 Step -1
    Dim J%: If Fnd_Idx(J, pAy, Rg(iRCnt, 1).Value) Then Stop: GoTo E
    If J = -1 Then Rg.Rows(iRCnt).EntireRow.Delete
Next
Exit Function
R: ss.R
E:
End Function

Function Dlt_TBar(pWs As Worksheet, pNmTBar$) As Boolean
Dim iOLEObj As Excel.OLEObject
For Each iOLEObj In pWs.OLEObjects
    If iOLEObj.Name = pNmTBar Then iOLEObj.Delete: Exit Function
Next
End Function

Function Dlt_Tbl(pNmt$, Optional pDb As database) As Boolean
Const cSub$ = "Dlt_Tbl"
Dim mDb As database: Set mDb = DbNz(pDb)
If Not IsTbl(pNmt, mDb) Then Exit Function
On Error GoTo R
If Left(pNmt, 1) = "[" And Right(pNmt, 1) = "]" Then
    mDb.TableDefs.Delete Mid(pNmt, 2, Len(pNmt) - 2)
Else
    mDb.TableDefs.Delete pNmt
End If
Exit Function
R: ss.R
E:
End Function

Function Dlt_Tbl_ByLnk() As Boolean
Const cSub$ = "Dlt_Tbl_ByLnk"
'Aim: Delete all linked table in currentdb
StsShw "Deleting all Link Tables  ..."
Dim mAnt_Lnk$(): If Fnd_Ant_ByLnk(mAnt_Lnk$) Then ss.A 1: GoTo E
Dim J%
For J = 0 To Sz(mAnt_Lnk) - 1
    If Dlt_Tbl(mAnt_Lnk(J)) Then ss.A 2: GoTo E
Next
GoTo X
R: ss.R
E:
X:
    Clr_Sts
End Function

Function Dlt_Tbl_ByPfx(pPfx$, Optional pDb As database) As Boolean
Const cSub$ = "Dlt_Tbl_ByPfx"
Dim mDb As database: Set mDb = DbNz(pDb)
Dim L%: L = Len(pPfx)
Dim mColl As New VBA.Collection
Dim iTbl As TableDef: For Each iTbl In mDb.TableDefs
    If Left(iTbl.Name, L) = pPfx Then mColl.Add iTbl.Name
Next
Dim mA$
Dim mNmt: For Each mNmt In mColl
    If Dlt_Tbl(CStr(mNmt), mDb) Then mA = Add_Str(mA, CStr(mNmt))
Next
mDb.TableDefs.Refresh
If Len(mA) <> 0 Then ss.A 1, "These tables cannot be deleted: " & mA: GoTo E
Exit Function
E:
End Function

Function Dlt_Ws(pWs As Worksheet) As Boolean
Const cSub$ = "Dlt_Ws_InWb"
On Error GoTo R
Dim mXls As Excel.Application: Set mXls = pWs.Application
mXls.DisplayAlerts = False
pWs.Delete
mXls.DisplayAlerts = True
Exit Function
R: ss.R
E:
End Function

Function Dlt_Ws_Excpt(pWb As Workbook, pWsNmExcpt$) As Boolean
'Aim: delete all ws except {pWsExcpt}
Const cSub$ = "Dlt_Ws_Excpt"
On Error GoTo R
pWb.Application.DisplayAlerts = False
While pWb.Sheets.Count >= 2
    If pWb.Sheets(1).Name = pWsNmExcpt Then
        pWb.Sheets(2).Delete
    Else
        pWb.Sheets(1).Delete
    End If
Wend
pWb.Application.DisplayAlerts = True
Exit Function
R: ss.R
E:
End Function

Function Dlt_Ws_Excpt__Tst()
Dim mWb As Workbook: If Crt_Wb(mWb, "c:\aa.xls", True) Then Stop
mWb.Sheets.Add
mWb.Sheets.Add
mWb.Sheets.Add
mWb.Application.Visible = True
Stop
If Dlt_Ws_Excpt(mWb, "ToBeDelete") Then Stop
Stop
mWb.Close True
End Function

Function Dlt_Ws_InWb(pWb As Workbook, pWsNm$) As Boolean
Const cSub$ = "Dlt_Ws_InWb"
On Error GoTo R
If Dlt_Ws(pWb.Worksheets(pWsNm)) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E:
End Function
