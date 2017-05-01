Attribute VB_Name = "nWrd_Fdoc"
Option Compare Database
Option Explicit

Function DocNew(Fdoc) As Word.Document
Dim O As Word.Document
Set O = Appw.Documents.Add
If Fdoc <> "" Then O.SaveAs Fdoc
End Function

Function FdocOpn(Fdoc, Optional Vis As Boolean) As Document
FfnAsstExist Fdoc, "FdocOpn"
Set FdocOpn = Appw(Vis).Documents.Open(Fdoc)
End Function

Sub FdocRpl(Fdoc$, RsHdr As DAO.Recordset, Optional RsDet As DAO.Recordset, Optional pNmDet$, Optional FfnDetTp$, Optional NHdrRows As Byte = 2)
'Aim: Substitue the [variables] in {pFfnDoc}.  The variables are in format of {xxx} where xxx is the fields of the {pRsHdr} or {pRsDet}.
'     {pRsDet} are always fill in "Word's Table" having substring {<<pNmDet>>} in cell(1,1).  Each record in will be filled starting from 3rd row of the table.
'     The row of the "Word's Table" will be created automatically
Const cSub$ = "Repl_Wrd"
Dim mWrd As Word.Document: If Opn_Wrd_RW(mWrd, Fdoc) Then ss.A 1: GoTo E
Dim iFld As DAO.Field
Dim mFnd As Word.Find: Set mFnd = mWrd.Range.Find

'With mFnd
'    .Forward = False
'    .ClearFormatting
'    .MatchWholeWord = False
'    .MatchCase = False
'    .Wrap = wdFindContinue
'End With
gWrd.ActiveWindow.ActivePane.View.Type = wdPrintView
'
'gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekPrimaryHeader
'For Each iFld In pRsHdr.Fields
'    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
'Next
'gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekPrimaryFooter
'For Each iFld In pRsHdr.Fields
'    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
'Next
'gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter
'For Each iFld In pRsHdr.Fields
'    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
'Next
'gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader
'For Each iFld In pRsHdr.Fields
'    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
'Next
gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument
For Each iFld In RsHdr.Fields
    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
Next

'-- Find if Detail Table exist ---------
If pNmDet = "" Then GoTo NoDet
If IsNothing(RsDet) Then GoTo NoDet

Dim mCase As Byte
mCase = 2
Select Case mCase
Case 1
    Dim iTbl As Word.Table, mFound As Boolean
    For Each iTbl In mWrd.Tables
        If iTbl.Rows.Count <> 3 Then GoTo NxtTbl
        If iTbl.Rows(1).Cells.Count <= 0 Then GoTo NxtTbl
        If InStr(iTbl.Rows(1).Cells(1).Range.Text, "{" & pNmDet & "}") = 0 Then GoTo NxtTbl
        mFound = True: Exit For
NxtTbl:
    Next
    If Not mFound Then GoTo NoDet
    '-- Replace {<<pNmDet>>} to empty
    mFnd.Execute "{" & pNmDet & "}", False, False, , , , False, , , "", WdReplace.wdReplaceAll
    '-- Detail ---------
    With RsDet
        iTbl.Rows(3).Select
        mWrd.Application.Selection.Copy
        While Not .EOF
            mWrd.Application.Selection.Paste
            .MoveNext
        Wend
        iTbl.Rows(3).Delete
        .MoveFirst

        Dim iRec%: iRec = 0
        While Not .EOF
            For Each iFld In RsDet.Fields
                With iTbl.Rows(3 + iRec).Range.Find
                    .Forward = False
                    .ClearFormatting
                    .MatchWholeWord = False
                    .MatchCase = False
                    .Wrap = wdFindStop
                    .Execute "{" & iFld.Name & "}", , , , , , , , , Nz(iFld.Value, ""), WdReplace.wdReplaceOne
                End With
            Next
            iRec = iRec + 1
            .MoveNext
        Wend
    End With
Case 2
    If VBA.Dir(FfnDetTp) = "" Then ss.A 3, "Template file for Detail Records does not exist": GoTo E
    Dim mWb As Workbook ' The Tp WB needs to keep open so that the format can be copied from source clip board
    '
    Stop
    'If Crt_Clip_ByRs(pFfnDetTp$, 3, pRsDet, mWb) Then ss.A 2:Goto E
    With mWrd.Application.Selection.Find
        .ClearFormatting
        .Text = "{" & pNmDet & "}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        If .Execute Then mWrd.Application.Selection.Paste
        Cls_Wb mWb
    End With
    'Assume there is only one table
    Dim iRow%
    For iRow = 1 To NHdrRows
        mWrd.Tables(1).Rows(iRow).HeadingFormat = True
    Next
End Select

NoDet:
  '  If FdocCls(mWrd, True) Then ss.A 3: GoTo E
Exit Sub
R: ss.R
E:
End Sub

Function FdocRpl__Tst()
Const cFfn$ = "c:\aa.doc"
'Dim mFfnTp$: mFfnTp = "C:\DOC1.DOC"
Dim mFbOldQsTmp$: If Fnd_Sffn_LgcMdbTmp(mFbOldQsTmp, "GenRmd") Then Stop
If TblCrt_FmLnkLnt(mFbOldQsTmp, "tmpBldOneRmd_Hdr,tmpBldOneRmd_Det") Then Stop
Dim mFfnTp$: mFfnTp = "M:\07 ARCollection\ARCollection\WorkingDir\Templates\Template_ReminderLvl3(English).doc"
Dim mRsHdr As DAO.Recordset: Set mRsHdr = CurrentDb.TableDefs("tmpBldOneRmd_Hdr").OpenRecordset
Dim mRsDet As DAO.Recordset: Set mRsDet = CurrentDb.TableDefs("tmpBldOneRmd_Det").OpenRecordset
If Cpy_Fil(mFfnTp, cFfn) Then Stop
If Repl_Wrd(cFfn, mRsHdr, mRsDet, "InvDet", Sffn_Tp("RmdInvDet(English)")) Then Stop
gWrd.Documents.Open cFfn
gWrd.Visible = True
End Function

Function FdocToStr$(A As Word.Document)
On Error GoTo R
FdocToStr = A.FullName
Exit Function
R: ss.R
    FdocToStr = ErStr("FdocToStr")
End Function

Sub FdocWrtPdf(Fdoc$, Optional Fpdf$, Optional KeepDocx As Boolean)
Dim W As Word.Document: Set W = FdocOpn(Fdoc)
'DocWrtPdf W
'DocCls W, NoSav:=True
If Not KeepDocx Then FfnDlt Fdoc
End Sub

Sub FdocWrtPdf__Tst()
FfnDlt "c:\RmdLvl1.Pdf": FfnDlt "c:\RmdLvl1.doc": If Cpy_Fil(Sffn_Tp("ReminderLvl1(English)", , ".doc"), "c:\RmdLvl1.doc") Then Stop: GoTo E
FfnDlt "c:\RmdLvl2.Pdf": FfnDlt "c:\RmdLvl2.doc": If Cpy_Fil(Sffn_Tp("ReminderLvl2(English)", , ".doc"), "c:\RmdLvl2.doc") Then Stop: GoTo E
FfnDlt "c:\RmdLvl3.Pdf": FfnDlt "c:\RmdLvl3.doc": If Cpy_Fil(Sffn_Tp("ReminderLvl3(English)", , ".doc"), "c:\RmdLvl3.doc") Then Stop: GoTo E
If Crt_PDF_FmWrd("c:\RmdLvl1.doc") Then GoTo E
If Crt_PDF_FmWrd("c:\RmdLvl2.doc") Then GoTo E
If Crt_PDF_FmWrd("c:\RmdLvl3.doc") Then GoTo E
If Opn_PDF("c:\RmdLvl1.pdf") Then ss.A 1: GoTo E
If Opn_PDF("c:\RmdLvl2.pdf") Then ss.A 2: GoTo E
If Opn_PDF("c:\RmdLvl3.pdf") Then ss.A 3: GoTo E
Exit Sub
R: ss.R
E:
End Sub
