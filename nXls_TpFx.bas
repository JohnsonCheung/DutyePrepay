Attribute VB_Name = "nXls_TpFx"
Option Compare Text
Option Explicit
Option Base 0

Sub TpFxClr()
Const cSub$ = "Clr_Tp"
Dim mDirTp$: mDirTp = Sdir_Tp
'-- Delete all tmp file records & insert one dummy record for those tmp*Output*
Dim iTbl As DAO.TableDef: For Each iTbl In CurrentDb.TableDefs
    If Left(iTbl.Name, 3) = "tmp" Then
        Debug.Print ": Deleting tmp table -->" & iTbl.Name
        If Run_Sql("Delete * from [" & iTbl.Name & "]") Then ss.A 1: GoTo E
        If InStr(iTbl.Name, "Output") > 0 Then
            With CurrentDb.TableDefs(iTbl.Name).OpenRecordset
                .AddNew
                .Update
            End With
        End If
    End If
Next
'Loop Template file
Dim AyFn$(): If Fnd_AyFn(AyFn, mDirTp, "*.xls", False) Then ss.A 1: GoTo E
If Sz(AyFn) = 0 Then ss.A 1, "No template files found", eRunTimErr, "DirTp", mDirTp: GoTo E
Dim iFil
For Each iFil In AyFn
    Debug.Print "******************************"
    Debug.Print "******************************"
    Debug.Print "Open Xls: " & iFil
    Dim mWb As Workbook: If Opn_Wb_RW(mWb, mDirTp & iFil) Then Stop
    If Rfh_Wb(mWb) Then ss.A 1: GoTo E
    Cls_Wb mWb, True
Next
MsgBox "All Templates Cleared"
Exit Sub
R: ss.R
E:
End Sub

Sub TpFxWrtFt()
Dim mF As Byte, mOFil$
mOFil = Sdir_Doc & "Tp_Doc.csv"
If Opn_Fil_ForOutput(mF, mOFil, True) Then ss.A 1: GoTo E

Dim mDirTp$: mDirTp = Sdir_Tp
If TpFxWrtFt_InDir(mDirTp, mF) Then ss.A 2: GoTo E

Dim iSubFolder As Folder
For Each iSubFolder In G.gFso.GetFolder(mDirTp).SubFolders
    Dim mDir As String
    mDir = iSubFolder.Name
    If mDir <> "." And mDir <> ".." Then If TpFxWrtFt_InDir(mDirTp & mDir, mF) Then ss.A 3: GoTo E
Next
Close #mF
'Format the csv to xls
Dim mWb As Workbook, mWs As Worksheet
If Opn_Wb_RW(mWb, mOFil) Then ss.A 4: GoTo E
Set mWs = mWb.Worksheets(1)
If WsFmtOL(mWs, 3) Then ss.A 5: GoTo E
mWs.Columns(3).ColumnWidth = 40
mWs.Columns(4).ColumnWidth = 15
FfnDlt Left(mOFil, Len(mOFil) - 4) & ".xls"
mWb.SaveAs Left(mOFil, Len(mOFil) - 4) & ".xls", Excel.XlFileFormat.xlWorkbookNormal
mWb.Application.Visible = True
Exit Sub
R: ss.R
E:
End Sub

Function TpFxWrtFt_InDir(pDirTp$, pF As Byte) As Boolean
'Aim: Exp all the datasource of all xls files in {pDir} to <pF>
Const cSub$ = "TpFxWrtFt_InDir"
'==Start==
Dim mAyFn$(): If Fnd_AyFn(mAyFn, pDirTp) Then ss.A 1: GoTo E
Dim J%
For J = 0 To Sz(mAyFn) - 1
    Dim mWb As Workbook: If Opn_Wb_R(mWb, pDirTp & mAyFn(J)) Then ss.A 1: GoTo E
    With mWb
        Write #pF, mWb.Name, , , , mWb.FullName
        If mWb.PivotCaches.Count > 0 Then
            Write #pF, , "PivotCaches.Count(" & mWb.PivotCaches.Count & ")"
            Dim iPc As Excel.PivotCache
            For Each iPc In .PivotCaches
                Write #pF, , , iPc.CtCommandText, , iPc.Connection
            Next
        End If
        Dim iWs As Worksheet
        For Each iWs In mWb.Worksheets
            If iWs.PivotTables.Count > 0 Then
                Write #pF, , "PivotTables.Count(" & iWs.PivotTables.Count & ") Ws(" & iWs.Name & ")"
                Dim iPt As PivotTable
                For Each iPt In iWs.PivotTables
                    Write #pF, , , iPt.PivotCache.CtCommandText, iPt.Name, iPt.PivotCache.Connection
                Next
            End If
        Next
        For Each iWs In mWb.Worksheets
            If iWs.QueryTables.Count > 0 Then
                Write #pF, , "QueryTables.Count(" & iWs.QueryTables.Count & ") Ws(" & iWs.Name & ")"
                Dim iQt As Excel.QueryTable
                For Each iQt In iWs.QueryTables
                    Write #pF, , , iQt.CtCommandText, iQt.Name, iQt.Connection
                Next
            End If
        Next
        .Close False
    End With
Next
Exit Function
R: ss.R
E:
End Function
