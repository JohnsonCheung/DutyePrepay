Attribute VB_Name = "nXls_nAct_nFmt_Ws"
Option Compare Database
Option Explicit

Function WsFmt(Rg As Range, pNRec&, Optional pRnoColrIdx& = 0) As Boolean
Const cSub$ = "WsFmt"
Dim mStp!
On Error GoTo R
If Rg.Row >= 4 Then Rg(-2, 1).Value = "Export @ " & Now()

mStp = 1
'Set SubTotal
Dim mCnoLas As Byte: If Fnd_CnoLas(mCnoLas, Rg(0, 1)) Then ss.A 1: GoTo E
Dim mWs As Worksheet: Set mWs = Rg.Parent
With Rg(1 + pNRec + 1, 2)
    If pNRec = 0 Then
        .Formula = 0
    Else
        Dim mCno As Byte: mCno = Rg.Column + 1
        Dim mRge As Range: Set mRge = mWs.Range(Rg(1, mCno), Rg(pNRec, mCno))
        Dim mAdr$: mAdr = mRge.Address(False, False)
        .Formula = Q_S(mAdr, "=SUBTOTAL(3,*)")
    End If
    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
End With

mStp = 2
'Copy Formula
If pNRec > 0 Then If RgCpyFormulaByCmt(Rg, pNRec) Then ss.A 2: GoTo E

mStp = 3
'MergeCells: If cell "A<cRnoDta> has color, it means it need to merge cells
If WsFmt_ByMge(Rg, mCnoLas, pNRec, Rg.Row - 1) Then ss.A 3: GoTo E

mStp = 4
'Use Row as color index
If pNRec > 0 Then If pRnoColrIdx > 0 Then If WsFmt_ByColrRow(Rg, pRnoColrIdx) Then ss.A 4: GoTo E

With Rg.Parent
    'If Not pNoExpTim Then .Range("A2").Value = "Exported @ " & Format(Now(), "yyyy/mm/dd hh:mm:ss")
    .Select
    .Activate
    .Outline.ShowLevels , 1
End With
With Rg(1, 1)
    .Select
    .Activate
End With
Exit Function
R: ss.R
E:
End Function

Function WsFmt_ByColrRow(Rg As Range, pRnoColrIdx&) As Boolean
'Aim: Color some columns cell as indexed by {pRnoColrIdx}
'     Which column to Colr: Find mAyCno() & mAyColr() @ {pRnoColrIdx} if they have colour
'     Detect Last Column  : use also non-empty-cell @ cRnoNmFld
'     Detect Last Row     : non-empty-cell of column B
'     Starting Row        : gRnoDta
Const cSub$ = "WsFmt_ByColrRow"
'On Error GoTo R
'If pRnoColrIdx <= 0 Then Exit Function
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'Shw_AllDta mWs
'Dim iCno As Byte, mSq As New cSq
'With mSq
'    .Rno1 = Rg.Row
'    .Rno2 = Rg.End(xlDown).Row
'End With
'
'Dim mAyCno() As Byte, mAyColr&()
'If Fnd_AyCnoColr(mAyCno, mAyColr, Rg, pRnoColrIdx) Then ss.A 1: GoTo E
'Dim J%
'For J = 0 To Sz(mAyCno) - 1
'    With mSq
'        .Cno1 = mAyCno(J)
'        .Cno2 = .Cno1
'    End With
'    Dim mRge As Range: If mSq.GetRge(mRge, mWs) Then ss.A 2: GoTo E
'    mRge.Interior.Color = mAyColr(J)
'Next
Exit Function
R: ss.R
E:
End Function

Function WsFmt_ByColrRow__Tst()
If Cpy_Fil("P:\AppDef_Meta\MetaLgc.xls", "c:\Tmp\aa.xls", True) Then Stop: GoTo E
Dim mWb As Workbook: If Opn_Wb_RW(mWb, "c:\tmp\aa.xls", , True) Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets("Stp")
If WsFmt_ByColrRow(mWs.Range("A5"), 3) Then Stop: GoTo E
Stop
GoTo X
E:
X: Cls_Wb mWb, False, True
End Function

Function WsFmt_ByMge(Rg As Range, pCnoEnd As Byte, pNRec&, pRnoMgeIdx As Byte) As Boolean
'Aim: Mge some columns cell as indexed by {pRnoColrIdx}
'     Which column to Mge: Test the first cell of {pRnoMgeIdx} has color or not.
'                          If no color, no merge
'                          If Yes, use Col A as Key to merge cells.
'                          The columns to be merge is same color as the first cell. (mAyCno(mN%))
'     Detect Last Column : use also non-empty-cell @ cRnoNmFld
'     Detect Last Row    : non-empty-cell of column A
'     Starting Row       : gRnoDta
Const cSub$ = "WsFmt_ByMege"
On Error GoTo R
Dim mWs As Worksheet: Set mWs = Rg.Parent
Dim mColr&: mColr = mWs.Cells(pRnoMgeIdx, Rg.Column).Interior.Color
If mColr = CtColrNone Then Exit Function
Dim iCno As Byte, mAyCno() As Byte, mN%
For iCno = Rg.Column To pCnoEnd
    Dim mRge As Range: Set mRge = mWs.Cells(pRnoMgeIdx, iCno)
    If mRge.Interior.Color = mColr Then
        ReDim Preserve mAyCno(mN): mAyCno(mN) = iCno: mN = mN + 1
    End If
Next
Dim iRno&, mRnoBeg&
mRnoBeg = Rg.Row
Dim mCnoLookAt As Byte: mCnoLookAt = Rg.Column
Dim mV, mVLas: mVLas = mWs.Cells(mRnoBeg, mCnoLookAt)
For iRno = mRnoBeg To mRnoBeg + pNRec - 1
    If iRno Mod 50 = 0 Then StsShw "Merging cell at row " & iRno & "..."
    mV = mWs.Cells(iRno, mCnoLookAt).Value
    If mV <> mVLas Then
        If WsFmt_MgeCells(mWs, mRnoBeg, iRno - 1, mAyCno) Then ss.A 2: GoTo E
        mRnoBeg = iRno
        mVLas = mV
    End If
Next
GoTo X
R: ss.R
E:
X: Clr_Sts
End Function

Function WsFmt_ByMge__Tst()
If Cpy_Fil("P:\AppDef_Meta\MetaLgc.xls", "c:\Tmp\aa.xls", True) Then Stop: GoTo E
Dim mWb As Workbook: If Opn_Wb_RW(mWb, "c:\tmp\aa.xls", , True) Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets("OldQsT")
Dim mRge As Range: Set mRge = mWs.Range("C4")
Dim mCnoLas As Byte: If Fnd_CnoLas(mCnoLas, mRge) Then Stop: GoTo E
Dim mRnoLas&: If Fnd_RnoLas(mRnoLas, mRge) Then Stop: GoTo E
Dim mNRec&: mNRec = mRnoLas - 5
If WsFmt_ByMge(mRge, mCnoLas, mNRec, 4) Then Stop: GoTo E
Exit Function
E:
End Function

Function WsFmt_MgeCells(pWs As Worksheet, pRnoBeg&, pRnoEnd&, pAyCno() As Byte) As Boolean
'Aim: for each {pCno} merge cells vertically from {pRnoBeg} to {pRnoEnd}
Const cSub$ = "WsFmt_MgeCells"
If pRnoBeg >= pRnoEnd Then Exit Function

On Error GoTo R
Dim iCno As Byte, J%, mRge As Range
Dim mXls As Excel.Application: Set mXls = pWs.Application: mXls.ScreenUpdating = False: mXls.DisplayAlerts = False
On Error GoTo R
For J = 0 To Sz(pAyCno) - 1
    Set mRge = pWs.Range(pWs.Cells(pRnoBeg, pAyCno(J)), pWs.Cells(pRnoEnd, pAyCno(J)))
    With mRge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .MergeCells = False
        .Merge
    End With
Next
GoTo X
R: ss.R
E:
X: mXls.ScreenUpdating = True
   mXls.DisplayAlerts = True
End Function

Function WsFmtOL(pWs As Worksheet, pUpToLvl As Byte) As Boolean
'Aim: Use column 1-{pUpToLvl} to set the outline level of {pWs}
Const cSub$ = "WsFmtOL"
Dim iRno&, iCno As Byte
For iRno = 1 To 65536
    Dim mIsSet As Boolean:   mIsSet = False ' Is row set the level? If not set, then exit the iRno loop
    For iCno = 1 To pUpToLvl
        If Not IsEmpty(pWs.Cells(iRno, iCno)) Then
            If iCno > 1 Then pWs.Rows(iRno).OutlineLevel = iCno
            mIsSet = True
            Exit For
        End If
    Next
    If Not mIsSet Then Exit For
Next
With pWs
    With .Outline
        .SummaryRow = xlSummaryAbove
        .ShowLevels 2
    End With
    .Range("$1:$" & pUpToLvl - 1).ColumnWidth = 5
    .Activate
    .Application.ActiveWindow.Zoom = 85
End With
End Function

Sub WsFmtOL_ByCol(Rg As Range, Optional pCithOL As Byte = 1, Optional pCithIns As Byte = 0)
'Aim: Set outline level of data at {Rg} by the cells content.  Assume the cell content is outline level
'     Optionally insert cell and shift right at pCnoIns relative to Rg
'Const cSub$ = "WsFmtOL_ByCol"
'On Error GoTo R
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'Dim iRno&, mLvl As Byte, mV
'
'Dim mRnoLas&: If Fnd_RnoLas(mRnoLas, Rg) Then ss.A 1: GoTo E
'If pCithIns > 0 Then
'    Dim mCnoLas As Byte: mCnoLas = Rg(0, 1).End(xlToRight).Column
'    Dim mRgeBlk As Range: Set mRgeBlk = Rg.Range(Rg(1, 1), mWs.Cells(mRnoLas, mCnoLas))
'    mRgeBlk.Sort Key1:=Rg(1, pCithOL), Order1:=xlAscending, Header:=xlNo _
'        , MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
'
'    Dim mAyRgeRno() As tRgeRno: If Fnd_AyRgeRno(mAyRgeRno, Rg(1, pCithOL)) Then ss.A 3: GoTo E
'    Dim J%
'    Dim mCnoOL As Byte: mCnoOL = Rg.Column + pCithOL - 1
'    Dim mCnoIns As Byte: mCnoIns = Rg.Column + pCithIns - 1
'    For J = 0 To SzRgeRno(mAyRgeRno) - 1
'        With mAyRgeRno(J)
'            mV = mWs.Cells(.Fm, mCnoOL).Value
'            If VarType(mV) <> vbDouble Then ss.A 4, "The data type of mAyRgeRno(J).Fm content is not Double, which is used as OutLine", , "mAyRgeRno(J).Fm", .Fm: GoTo E
'            If 0 > mV Or mV > 15 Then ss.A 5, "The value of mAyRgeRno(J).Fm content is not between 0 to 15", , "mAyRgeRno(J).Fm,The Val", .Fm, mV: GoTo E
'            mLvl = mV
'            If mLvl >= 1 Then
'                Dim mRge As Range: Set mRge = mWs.Range(mWs.Cells(.Fm, mCnoIns), mWs.Cells(.To, mCnoIns + mLvl - 1))
'                mRge.Insert shift:=Excel.xlShiftToRight
'            End If
'        End With
'    Next
'
'    If Crt_Rge_ExtNCol(mRgeBlk, mRgeBlk, 15) Then ss.A 6: GoTo E
'    mRgeBlk.Sort Key1:=Rg(1, 1), Order1:=xlAscending, Header:=xlNo _
'        , MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
'End If
'
'For iRno = Rg.Row To mRnoLas
'    mLvl = mWs.Cells(iRno, mCnoOL).Value
'    If 1 <= mLvl And mLvl <= 7 Then mWs.Rows(iRno).OutlineLevel = mLvl + 1
'    If 8 <= mLvl And mLvl <= 15 Then mWs.Rows(iRno).OutlineLevel = 8
'Next
'
'With mWs
'    With .Outline
'        .SummaryRow = xlSummaryAbove
'        .ShowLevels 2
'    End With
'    .Activate
'    .Application.ActiveWindow.Zoom = 85
'End With
'Rg(1, pCithOL).EntireColumn.ColumnWidth = 5
'Exit Sub
'R: ss.R
'E: WsFmtOL_ByCol = True: ss.B cSub, cMod, "Rg,pCithOL,pCithIns", RgToStr(Rg), pCithOL, pCithIns
End Sub

Sub WsFmtOutLine(A As Worksheet)  '(P As WsFmtOL)
'Aim: Color, Outline & Border by {p}
'     Assume Col A is OutLine Level#.
Const cSub$ = "WsFmtOutLine"
'On Error GoTo R
''Find mAyColr() by p.NLvl & p.AyLvlCols
'ReDim mAyColr&(P.NLvl - 1)
'Dim iLvl%
'For iLvl = 0 To P.NLvl - 1
'    Dim mA$: mA = P.AyLvlCols(iLvl)
'    Dim mP%: mP% = InStr(mA, ":")
'    If mP > 0 Then mA = Left(mA, mP - 1)
'    Dim mRge As Range: Set mRge = P.Ws.Cells(P.RnoColr, mA)
'    mAyColr(iLvl) = mRge.Interior.Color
'Next
''
'Dim iRno&
'Dim mLvlLas%: mLvlLas = 0
'With P.Ws
'    For iRno = P.RnoFm To P.RnoTo
'        'set Outline
'        Dim mOutLine As Byte: mOutLine = .Cells(iRno, 1).Value
'        If mOutLine > 1 Then
'            Set mRge = .Rows(iRno)
'            mRge.EntireRow.OutlineLevel = mOutLine
'        End If
'
'        If mOutLine = mLvlLas Then GoTo Nx
'        mLvlLas = mOutLine
'
'        'Find 3 Ranges
'        Dim mRgeBorder As Range, mRgeColrRow As Range, mRgeColrCols As Range
'        Dim mLvlCur%: mLvlCur = .Cells(iRno, 1).Value
'        Dim mLvlCols$: mLvlCols = P.AyLvlCols(mLvlCur - 1)
'        Dim mColr&: mColr = mAyColr(mLvlCur - 1)
'        If WsFmtOutLine_Fnd2Rge(mRgeBorder, mRgeColrCols, P.Ws, iRno, P.RnoTo, mLvlCols, P.ColEnd) Then ss.A 1: GoTo E
'        'Fmt the 2 ranges
'        mRgeBorder.BorderAround XlLineStyle.xlContinuous, XlBorderWeight.xlMedium
'        mRgeBorder.Interior.Color = mColr
'
'        If mOutLine <> P.NLvl Then
'            Dim mSq As tSq:
'            Stop
'            'If xCv.RgSq(mSq, mRgeColrCols) Then ss.A 2: GoTo E
'            If mSq.R2 > mSq.R1 Then
'                mSq.R1 = mSq.R1 + 1
'                If xCv.Cv_Sq2Rge(mRgeColrCols, P.Ws, mSq) Then ss.A 3: GoTo E
'                mRgeColrCols.Font.Color = mColr
'            End If
'        End If
'Nx: Next
'End With
'Exit Sub
'R: ss.R
'E:
End Sub

Function WsFmtr(A As Worksheet) As LoFmtr
WsFmtr = LoFmtr(WsSq(A))
End Function

Private Function WsFmtOutLine_Fnd2Rge(ByRef oRgeBorder As Range, ByRef oRgeColrCols _
    , pWs As Worksheet, pRno&, pRnoTo&, pLvlCols$, pColEnd$) As Boolean
'Aim: Find oRge by p*
'     Assume the Lvl Col is at column A
'     oRgeBorder:   The whole range should be border arround.  Range of Row = Current row to [mRnoEnd],  Range of Col = first col of {pLvlCols} up to pColEnd$
'                   [mRnoEnd] = some row below of lvl =< pLvl or {pRnoTo}
'     pRno: current rno
'     pRnoTo last row
'     pLvlCols in "C" or "C:D" format
'     pColEnd  in "CC" format
Const cSub$ = "WsFmtOutLine_Fnd2Rge"
On Error GoTo R
'Find mColBeg
Dim mLvlColBeg$, mLvlColEnd$
Dim mP%: mP = InStr(pLvlCols, ":")
If mP = 0 Then
    mLvlColBeg = pLvlCols
    mLvlColEnd = pLvlCols
Else
    mLvlColBeg = Left(pLvlCols, mP - 1)
    mLvlColEnd = Mid(pLvlCols, mP + 1)
End If
'Find mRnoEnd: a Row just before a the row with Lvl =< CurLvl or the pRnoTo
Dim mRnoEnd&
Dim mLvlCur%: mLvlCur = pWs.Cells(pRno, 1).Value
Dim mFound As Boolean: mFound = False
Dim mBrk As Boolean: mBrk = False
For mRnoEnd = pRno + 1 To pRnoTo
    Dim iLvl%: iLvl = pWs.Cells(mRnoEnd, 1).Value
    If mLvlCur = iLvl Then
        If mBrk Then mRnoEnd = mRnoEnd - 1: mFound = True: Exit For
        GoTo Nx
    End If
    mBrk = True
    If iLvl <= mLvlCur Then mRnoEnd = mRnoEnd - 1: mFound = True: Exit For
Nx:
Next
If Not mFound Then mRnoEnd = pRnoTo
'Set oRgeBorder
Set oRgeBorder = pWs.Range(mLvlColBeg & pRno & ":" & pColEnd & mRnoEnd)
Set oRgeColrCols = pWs.Range(mLvlColBeg & pRno & ":" & mLvlColEnd & mRnoEnd)
Exit Function
R: ss.R
E:
End Function

