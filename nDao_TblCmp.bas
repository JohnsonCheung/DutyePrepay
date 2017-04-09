Attribute VB_Name = "nDao_TblCmp"
Option Compare Database
Option Explicit

Function TblCmp(T1$, T2$, pLoCmpKey$, pLoCmV$) As Boolean
Const cSub$ = "TblCmp"
'Debug.Print TblCmp("tmpAsAt_F0311_1Os_Odbc", "tmpCurOs_F0311_CurOs_Odbc", "RPAN8,RPDCT,RPDOC", "OsBas=RPAAP,OsCur=RPFAP")
'Aim: Build 5 steps of queries to compare 2 tables:
Const Nmq10$ = "qryCmp_01_0_Lst"
Const Nmq11$ = "qryCmp_01_1_Fm_Below"
Const Nmq12$ = "qryCmp_01_2_Lst"
Const Nmq20$ = "qryCmp_02_0_SumA"
Const Nmq21$ = "qryCmp_02_1_FmA"
Const Nmq30$ = "qryCmp_03_0_SumB"
Const Nmq31$ = "qryCmp_03_1_FmB"
Const Nmq40$ = "qryCmp_04_0_Output"
Const Nmq41$ = "qryCmp_04_1_Fm_Lst_SumA_SumB"
Const Nmq50$ = "qryCmp_05_0_Det"
Const Nmq51$ = "qryCmp_05_1_Fm_Lst_A_B"
'1 Lst
Const mSql10$ = "SELECT * FROM tmpCmp_Lst;"
Const mSql11$ = "SELECT * INTO tmpCmp_Lst FROM qryCmp_01_2_Lst;"
Dim mSql12$: mSql12$ = "SELECT {0} FROM {1} UNION SELECT {2} FROM {3}"
''Select RPAN8 as [K1_RPAN8],RPDCT as [K2_RPDCT], RPDOC as [K3_RPDOC] from tmpAsAt_F0311_1Os_Odbc"
'' UNION Select RPAN8 as [K1_RPAN8],RPDCT as [K2_RPDCT], RPDOC as [K3_RPDOC] from tmpCurOs_F0311_CurOs_Odbc;

'2 SumA
Const mSql20$ = "SELECT * FROM tmpCmp_SumA;"
Dim mSql21$: mSql21$ = "SELECT DISTINCT {0}, Count(*) as A_Cnt, {1} INTO tmpCmp_SumA From [{2}] AS a Group by {3}"
''SELECT DISTINCT [RPAN8] AS K1_RPAN8, [RPDCT] AS K2_RPDCT, [RPDOC] AS K3_RPDOC, Count(1) AS A_Cnt, Sum(a.[OsBas]) AS [A1_OsBas], Sum(a.[OsCur]) AS [A2_OsCur] INTO tmpCmp_SumA
'' FROM tmpAsAt_F0311_1Os_Odbc AS a
'' GROUP BY a.[RPAN8], a.[RPDCT], a.[RPDOC];

'3 SumB
Const mSql30$ = "SELECT * FROM tmpCmp_SumB;"
Dim mSql31$: mSql31$ = "SELECT DISTINCT {0}, Count(*) as B_Cnt, {1} INTO tmpCmp_SumB From [{2}] AS b Group by {3}"
''SELECT DISTINCT [RPAN8] AS K1_RPAN8, [RPDCT] AS K2_RPDCT, [RPDOC] AS K3_RPDOC, Count(1) AS B_Cnt, Sum(b.[RPAAP]) AS [B1_RPAAP], Sum(b.RPFAP) AS [B2_RPFAP] INTO tmpCmp_SumB
'' FROM tmpCurOs_F0311_CurOs_Odbc AS b
'' GROUP BY b.[RPAN8], b.[RPDCT], b.[RPDOC];

'4 Output
Const mSql40$ = "SELECT * FROM tmpCmp_Output;"
Dim mSql41$: mSql41$ = "SELECT {0}, A_Cnt, B_Cnt, {1}, {2} AS IsSam INTO tmpCmp_Output From (tmpCmp_Lst AS l LEFT JOIN tmpCmp_SumA AS a ON {3}) LEFT JOIN tmpCmp_SumB AS b ON {4}"
''SELECT l.K1_RPAN8, l.K2_RPDCT, l.K3_RPDOC, A_Cnt, B_Cnt, a.A1_OsBas, B1_RPAAP, a.A2_OsCur, B2_RPFAP, Nz([A1_OsBas],0)=Nz([B1_RPAAP],0) And Nz([A2_OsCur],0)=Nz([B2_RPFAP],0) AS IsSam INTO tmpCmp_Output
''FROM (tmpCmp_Lst AS l LEFT JOIN tmpCmp_SumA AS a ON (l.K3_RPDOC = a.K3_RPDOC) AND (l.K2_RPDCT = a.K2_RPDCT) AND (l.K1_RPAN8 = a.K1_RPAN8)) LEFT JOIN tmpCmp_SumB AS b ON (l.K3_RPDOC = K3_RPDOC) AND (l.K2_RPDCT = K2_RPDCT) AND (l.K1_RPAN8 = K1_RPAN8);

'5 Det
Const mSql50$ = "SELECT * FROM tmpCmp_Det"
Dim mSql51$: mSql51$ = "SELECT l.*, a.*, b.* from (tmpCmp_Lst As l Left Join [{0}] a on {1}) Left Join [{2}] b on {3}"
''SELECT l.*, a.*, b.*
''FROM (tmpCmp_Lst AS l left JOIN tmpAsAt_F0311_1Os_Odbc AS a ON (l.K3_RPDOC = a.RPDOC) AND (l.K2_RPDCT = a.RPDCT) AND (l.K1_RPAN8 = a.RPAN8)) left JOIN tmpCurOs_F0311_CurOs_Odbc AS b ON (l.K3_RPDOC = RPDOC) AND (l.K2_RPDCT = RPDCT) AND (l.K1_RPAN8 = RPAN8);

'Build Common Element
Dim mAm_NmFldKey() As tMap: mAm_NmFldKey = Get_Am_ByLm(pLoCmpKey)
Dim mAm_NmFldVal() As tMap: mAm_NmFldVal = Get_Am_ByLm(pLoCmV)
Dim NKey As Byte: NKey = Siz_Am(mAm_NmFldKey)
Dim NVal As Byte: NVal = Siz_Am(mAm_NmFldVal)
ReDim mAnFld_CmnKey$(0 To NKey - 1)
Dim J%: For J% = 0 To NKey - 1
    With mAm_NmFldKey(J%)
        mAnFld_CmnKey(J%) = "[K" & J% & "_" & IIf(.F1 = .F2, .F1, .F1 & "_" & .F2) & "]"
    End With
Next
'Set mSql*$
ReDim mAm(0 To NKey - 1) As tMap
Dim A0$, A1$, A2$, A3$, A4$
''=====Dim mSql12$: mSql12$ = "SELECT {0} FROM {1} UNION SELECT {2} FROM {3}"
'''Select RPAN8 as [K1_RPAN8],RPDCT as [K2_RPDCT], RPDOC as [K3_RPDOC] from tmpAsAt_F0311_1Os_Odbc"
''' UNION Select RPAN8 as [K1_RPAN8],RPDCT as [K2_RPDCT], RPDOC as [K3_RPDOC] from tmpCurOs_F0311_CurOs_Odbc;
''''A0, A1
If Cpy_Am(mAm, mAm_NmFldKey) Then ss.A 3: GoTo E
If Set_Am_F2(mAm, mAnFld_CmnKey) Then ss.A 4: GoTo E
A0 = ToStr_Am(mAm, " AS ", "[]")
A1 = T1

''''A2, A3
If Cpy_Am(mAm, mAm_NmFldKey, True) Then ss.A 5: GoTo E
If Set_Am_F2(mAm, mAnFld_CmnKey) Then ss.A 6: GoTo E
A2 = ToStr_Am(mAm, " AS ", "[]")
A3 = T2

''''mSql12
mSql12 = Fmt_Str(mSql12, A0, A1, A2, A3)

''=====Dim mSql21$: mSql21$ = "SELECT DISTINCT {0}, Count(*) as A_Cnt, {1} INTO tmpCmp_SumA From {2} AS a Group by {3}"
'''SELECT DISTINCT [RPAN8] AS K1_RPAN8, [RPDCT] AS K2_RPDCT, [RPDOC] AS K3_RPDOC, Count(1) AS A_Cnt, Sum(a.[OsBas]) AS [A1_OsBas], Sum(a.[OsCur]) AS [A2_OsCur] INTO tmpCmp_SumA
''' FROM tmpAsAt_F0311_1Os_Odbc AS a
''' GROUP BY a.[RPAN8], a.[RPDCT], a.[RPDOC];
''''A0
If Cpy_Am(mAm, mAm_NmFldKey) Then ss.A 7: GoTo E
If Set_Am_F2(mAm, mAnFld_CmnKey) Then ss.A 8: GoTo E
A0 = ToStr_Am(mAm, " AS ", "[]")

''''A1
If Cpy_Am(mAm, mAm_NmFldVal) Then ss.A 9: GoTo E
If Cpy_Am_F1ToF2(mAm) Then ss.A 10: GoTo E
A1 = ToStr_Am(mAm, " AS ", "Sum(a.[*])", "[A{N}_*]")

''''A2
A2 = T1

''''A3
A3 = ToStr_AmF1(mAm_NmFldKey, "a.[*]")

''''mSql21
mSql21 = Fmt_Str(mSql21, A0, A1, A2, A3)

''=====Dim mSql31$: mSql31$ = "SELECT DISTINCT {0}, Count(*) as B_Cnt, {1} INTO tmpCmp_SumB From [{2}] AS b Group by {3}"
'''SELECT DISTINCT [RPAN8] AS K1_RPAN8, [RPDCT] AS K2_RPDCT, [RPDOC] AS K3_RPDOC, Count(1) AS B_Cnt, Sum(b.[RPAAP]) AS [B1_RPAAP], Sum(b.RPFAP) AS [B2_RPFAP] INTO tmpCmp_SumB
''' FROM tmpCurOs_F0311_CurOs_Odbc AS b
''' GROUP BY b.[RPAN8], b.[RPDCT], b.[RPDOC];
''''A0
If Cpy_Am(mAm, mAm_NmFldKey, True) Then ss.A 11: GoTo E
If Set_Am_F2(mAm, mAnFld_CmnKey) Then ss.A 12: GoTo E
A0 = ToStr_Am(mAm, " AS ", "[*]")

''''A1
If Cpy_Am(mAm, mAm_NmFldVal) Then ss.A 13: GoTo E
If Cpy_Am_F2ToF1(mAm) Then ss.A 14: GoTo E
A1 = ToStr_Am(mAm, " AS ", "Sum(b.[*])", "[B{N}_*]")

''''A2
A2 = T2

''''A3
A3 = ToStr_AmF2(mAm_NmFldKey, "b.[*]")

''''mSql31
mSql31 = Fmt_Str(mSql31, A0, A1, A2, A3)

'''=====Dim mSql41$: mSql41$ = "SELECT {0}, A_Cnt, B_Cnt, {1}, {2} AS IsSam INTO tmpCmp_Output From (tmpCmp_Lst AS l LEFT JOIN tmpCmp_SumA AS a ON {3}) LEFT JOIN tmpCmp_SumB AS b ON {4}"
'''SELECT l.K1_RPAN8, l.K2_RPDCT, l.K3_RPDOC, A_Cnt, B_Cnt, [A1_OsBas], [B1_RPAAP], [A2_OsCur], [B2_RPFAP], Nz([A1_OsBas],0)=Nz([B1_RPAAP],0) And Nz([A2_OsCur],0)=Nz([B2_RPFAP],0) AS IsSam INTO tmpCmp_Output
'''FROM (tmpCmp_Lst AS l LEFT JOIN tmpCmp_SumA AS a ON (l.K3_RPDOC = a.K3_RPDOC) AND (l.K2_RPDCT = a.K2_RPDCT) AND (l.K1_RPAN8 = a.K1_RPAN8)) LEFT JOIN tmpCmp_SumB AS b ON (l.K3_RPDOC = K3_RPDOC) AND (l.K2_RPDCT = K2_RPDCT) AND (l.K1_RPAN8 = K1_RPAN8);
''''A0
A0 = ToStr_Ays(mAnFld_CmnKey, "l.*")

''''A1
A1 = ToStr_Am(mAm_NmFldVal, ", ", "[A{N}_*]", "[B{N}_*]")

''''A2
A2 = ToStr_Am(mAm_NmFldVal, "=", "Nz([A{N}_*],0)", "Nz([B{N}_*],0)", " and ")

''''A3
ReDim mAm(0 To Siz_Ay(mAnFld_CmnKey) - 1) As tMap
If Set_Am_F1(mAm, mAnFld_CmnKey) Then ss.A 15: GoTo E
If Set_Am_F2(mAm, mAnFld_CmnKey) Then ss.A 16: GoTo E
A3 = ToStr_Am(mAm, "=", "(l.*", "a.*)", " and ")

''''A4
A4 = ToStr_Am(mAm, "=", "(l.*", "b.*)", " and ")

''''mSql41
mSql41 = Fmt_Str(mSql41, A0, A1, A2, A3, A4)

''=====Dim mSql51$: mSql51$ = "SELECT l.*, a.*, * from (tmpCmp_Lst As l Left Join {0} on {1}) Left Join {2} on {3}"
'''SELECT l.*, a.*, *
'''FROM (tmpCmp_Lst AS l left JOIN tmpAsAt_F0311_1Os_Odbc AS a ON (l.[K3_RPDOC] = a.[RPDOC]) AND (l.[K2_RPDCT] = a.[RPDCT]) AND (l.[K1_RPAN8] = a.[RPAN8])) left JOIN tmpCurOs_F0311_CurOs_Odbc AS b ON (l.K3_RPDOC = RPDOC) AND (l.K2_RPDCT = RPDCT) AND (l.K1_RPAN8 = RPAN8);
''''A0,A1
A0 = T1

If Cpy_Am(mAm, mAm_NmFldKey, True) Then ss.A 17: GoTo E
If Set_Am_F1(mAm, mAnFld_CmnKey) Then ss.A 18: GoTo E
A1 = ToStr_Am(mAm, "=", "(l.*", "a.[*])", " and ")

''''A2,A3
If Cpy_Am(mAm, mAm_NmFldKey) Then ss.A 19: GoTo E
If Set_Am_F1(mAm, mAnFld_CmnKey) Then ss.A 20: GoTo E
A2 = T2
A3 = ToStr_Am(mAm, "=", "(l.*", "b.[*])", " and ")

mSql51 = Fmt_Str(mSql51, A0, A1, A2, A3, A4)
''====Create queries & Run
If QryCrt(Nmq10, mSql10) Then ss.A 21: GoTo E
If QryCrt(Nmq11, mSql11) Then ss.A 22: GoTo E
If QryCrt(Nmq12, mSql12) Then ss.A 23: GoTo E
If QryCrt(Nmq20, mSql20) Then ss.A 24: GoTo E
If QryCrt(Nmq21, mSql21) Then ss.A 25: GoTo E
If QryCrt(Nmq30, mSql30) Then ss.A 26: GoTo E
If QryCrt(Nmq31, mSql31) Then ss.A 27: GoTo E
If QryCrt(Nmq40, mSql40) Then ss.A 28: GoTo E
If QryCrt(Nmq41, mSql41) Then ss.A 29: GoTo E
If QryCrt(Nmq50, mSql50) Then ss.A 30: GoTo E
If QryCrt(Nmq51, mSql51) Then ss.A 31: GoTo E
If Run_Qry("qryCmp", , , True) Then ss.A 32: GoTo E

'Create relation
Dim mAy$(): If Cpy_AmF1_ToAy(mAy, mAm_NmFldKey) Then ss.A 33: GoTo E
If Crt_TqRel("qryCmp_04_0_Output", T2, ToStr_Ays(mAnFld_CmnKey, , ";"), ToStr_Ays(mAy, , ";")) Then ss.A 34: GoTo E

'OpnQry
If Opn_Qry(Nmq40) Then ss.A 35: GoTo E
Exit Function
R: ss.R
E: TblCmp = True: ss.B cSub, cMod, "T1,T2,pLoCmpKey,pLoCmV", T1, T2, pLoCmpKey, pLoCmV
End Function

Function TblCmp__Tst()
'Debug.Print TblCmp("tmpChk_Hdr", "qF0311", "RPAN8,RPDCT,RPDOC,RPSFX,RPCRCD", "RPAAP=RPAG,RPFAP=RPACR")
'Debug.Print TblCmp("tmpARBalAt_F0311_1At_Odbc", "tmpARBalCur_F0311_1Cur_Odbc", "RPAN8,RPDCT,RPDOC,RPCRCD", "OsBas=RPAAP,OsCur=RPFAP")
Debug.Print TblCmp("mstBrand", "mstBrand", "Brand", "BrandId")
End Function

Function TblCmp_x(TBef$, TAft$, pLnKey$, FnStr$) As Boolean
Const cSub$ = "TblCmp_x"
'Aim: Update <TAft> table: <IsChg>, <Refreshed>, <Rmk>, <<FnStr>>
'Assume: There is <IsChg>, <Refreshed>, <Rmk> field in <TAft> table
'TBef: Table name of before image
'TAft: Table name of after image
'pLnKey: Key list used to join in format: K1,Kb2/Ka2,..    K1 is common field name, Kb2/Ka2 is pair of key field of aft and bef of different name
'FnStr: Field List required to compare, same format as above
'==Start
'Break pLnKey and FnStr into AyK(), bKey() / aFld(), bFld()
Dim Am1() As tMap, Am2() As tMap
Am1 = Get_Am_ByLm(pLnKey)
Am2 = Get_Am_ByLm(FnStr)

''Build open Rs Sql String
'Dim mBefLst$
'Dim mAftLst$
'Dim mJoin$
'ReDim mAy$(LBound(aFld) To UBound(aFld))
'Dim J As Byte
'For J = LBound(aFld) To UBound(aFld)
'    mAy(J) = "a.[" & aFld(J) & "] As [a_" & aFld(J) & "]"
'Next
'mAftLst = Join(mAy, ", ")
'For J = LBound(bFld) To UBound(bFld)
'    mAy(J) = "[" & bFld(J) & "] As [b_" & bFld(J) & "]"
'Next
'mBefLst = Join(mAy, ", ")
'ReDim mAy$(LBound(AyK) To UBound(AyK))
'For J = LBound(AyK) To UBound(AyK)
'    mAy(J) = "a.[" & AyK(J) & "]=[" & bKey(J) & "]"
'Next
'mJoin = Join(mAy, " and ")
'
'mSql = Fmt_Str( _
'    "Select a.IsChg as a_IsChg, a.Rmk as a_Rmk, a.Refreshed as a_Refreshed, {0}, {1}" & _
'    " from [{2}] b" & _
'    " inner join [{3}] a on {4}", _
'        mBefLst, mAftLst, TBef, TAft, mJoin)
'Stop
''If pIsChk Then
''    ShowDbgPrompt "Sql Str to open bef/aft table to compare", "Sql Str"
''    debug.print  "BefLst=" & mBefLst
''    debug.print  "AftLst=" & mAftLst
''    debug.print  "TBef(table)=" & TBef
''    debug.print  "TAft(table)=" & TAft
''    debug.print  "Join=" & mJoin
''    debug.print  mSql
''    Stop
''End If
''Open Rs and Loop
'With CurrentDb.OpenRecordset(mSql)
'    While Not .EOF
'        ''If Not equal
'        Dim mIsSam As Boolean
'        mIsSam = True
'        For J = LBound(aFld) To UBound(aFld)
'            If .Fields("b_" & bFld(J)).Value <> .Fields("a_" & aFld(J)).Value Then
'                mIsSam = False
'                Exit For
'            End If
'        Next
'
'        Stop
''        If pIsChk Then
''            ShowDbgPrompt "Whether Bef/Aft table is same or not?", "Is Same?"
''            For J = LBound(aFld) To UBound(aFld)
''                If bFld(J) = aFld(J) Then
''                    debug.print  aFld(J); "=";
''                Else
''                    debug.print  bFld(J); "/"; aFld(J); "=";
''                End If
''                debug.print  "[" & .Fields("b_" & bFld(J)).Value & "]/[" & .Fields("a_" & aFld(J)).Value & "]"
''            Next
''            debug.print  "mIsSam=" & mIsSam
''            Stop
''        End If
'        If Not mIsSam Then
'            '''Build <Rmk> & <mSetLst>
'            Dim mSetLst$, mRmk$
'            mRmk = ""
'            For J = LBound(aFld) To UBound(aFld)
'                If bFld(J) = aFld(J) Then
'                    mRmk = mRmk & aFld(J) & "="
'                Else
'                    mRmk = mRmk & bFld(J) & "/" & aFld(J) & "="
'                End If
'                Dim mV$
'                If .Fields("b_" & bFld(J)).Value = .Fields("a_" & aFld(J)).Value Then
'                    mV = "[" & .Fields("b_" & bFld(J)).Value & "]"
'                Else
'                    mV = "[" & .Fields("b_" & bFld(J)).Value & "]/[" & .Fields("a_" & aFld(J)).Value & "]<--"
'                End If
'                mRmk = mRmk & mV & vbCrLf
'            Next
'            mRmk = mRmk & "----" & vbCrLf
'
'            Stop
''            If pIsChk Then
''                ShowDbgPrompt "Remark", "Check Remark"
''                For J = LBound(aFld) To UBound(aFld)
''                    If bFld(J) = aFld(J) Then
''                        debug.print  aFld(J); "=";
''                    Else
''                        debug.print  bFld(J); "/"; aFld(J); "=";
''                    End If
''                    debug.print  "[" & .Fields("b_" & bFld(J)).Value & "]/[" & .Fields("a_" & aFld(J)).Value & "]"
''                Next
''                debug.print  "mIsSam=" & mIsSam
''                Stop
''            End If
'
'            '''Update the fields of <TAft>: <IsChg>, <Refreshed>, <Rmk>, <<FnStr>>
'            ''''Build mSql for Update "Update {0} Set {1} Where {2}
'
'            Dim mWhere$
'            ReDim mAy$(LBound(aFld) To UBound(aFld))
'            For J = LBound(aFld) To UBound(aFld)
'                mAy(J) = "a.[" & aFld(J) & "] As [a_" & aFld(J) & "]"
'            Next
'            mAftLst = Join(mAy, ", ")
'            For J = LBound(bFld) To UBound(bFld)
'                mAy(J) = "[" & bFld(J) & "] As [b_" & bFld(J) & "]"
'            Next
'            mBefLst = Join(mAy, ", ")
'            '''''mWhere
'            ReDim mAy$(LBound(AyK) To UBound(AyK))
'            For J = LBound(AyK) To UBound(AyK)
'                mAy(J) = "a.[" & AyK(J) & "]=[" & bKey(J) & "]"
'            Next
'            mJoin = Join(mAy, " and ")
'
'            mSql = Fmt_Str("Update [{0}] Set {1} Where {2}", _
'                    TAft, mSetLst, mWhere)
'            Stop
''            If pIsChk Then
''                ShowDbgPrompt "Sql Str to Update one record", "Sql Str"
''                debug.print  "TAft(table)=" & TAft
''                debug.print  "SetLst=" & mSetLst
''                debug.print  "Where=" & mWhere
''                debug.print  mSql
''                Stop
''            End If
'            SqlRun mSql
'
'            '.Edit
'            '!a_IsChg = True
'            '!a_Refreshed = Now
'            '!a_Rmk = !a_Rmk & mRmk
'            'For J = LBound(aFld) To UBound(aFld)
'            '    .Fields("a_" & aFld(J)).Value = .Fields("b_" & bFld(J)).Value
'            'Next
'            .Update
'        End If
'        .MoveNext
'    Wend
'    .Close
'End With
Exit Function
R: ss.R
E: TblCmp_x = True: ss.B cSub, cMod, "TBef$, TAft$, pLnKey$, FnStr$", TBef$, TAft$, pLnKey$, FnStr$
End Function

