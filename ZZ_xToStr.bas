Attribute VB_Name = "ZZ_xToStr"
Function AppaToStr$(A As Access.Application)
If TypeName(A) = "Nothing" Then AppaToStr = "Nothing": Exit Function
AppaToStr = A.Name
End Function
Function AppaToStr1$(A As Access.Application)
On Error GoTo R
If TypeName(A) = "Nothing" Then AppaToStr1 = "Nothing": Exit Function
AppaToStr 1 = pAcs.CurrentDb.Name
Exit Function
R:
AppaToStr1 = "AppaToStr error. Msg=[" & Err.Description & "]"
End Function

Function ToStr_PgmADcl$(pArgDcl$ _
    , Optional pIsOpt As Boolean = False, Optional pIsByVal As Boolean = False, Optional pDftVal$ = "")
Dim mA$: mA = IIf(pIsByVal, "ByVal ", "") & pArgDcl & Cv_Str(pDftVal, "=")
If pIsOpt Then ToStr_PgmADcl = Q_S(mA, "[]"): Exit Function
ToStr_PgmADcl = mA
End Function
Function ToStr_Sgnt$(pNmTypArg$ _
    , Optional pIsOpt As Boolean = False _
    , Optional pIsByVal As Boolean = False _
    , Optional pIsAy As Boolean = False _
    , Optional pIsPrmAy As Boolean _
    , Optional pDftVal$ = "" _
    )
Dim mA$: mA = Cv_Bool(pIsPrmAy, "PrmAy ") & pNmTypArg & Cv_Bool(pIsAy, "()") & Cv_Str(pDftVal, "=")
If pIsOpt Then ToStr_Sgnt = Q_S(mA, "[]"): Exit Function
ToStr_Sgnt = mA
End Function

Function ToStr_Sgnt__Tst()
Debug.Print ToStr_Sgnt("sdlkfj", True, True, True, True, "skdjf")
Debug.Print ToStr_Sgnt("sdlkfj")
End Function

Function ToStr_ArgDcl$(pNmArg$ _
    , Optional pNmTypArg$ = "$", Optional pIsPrmAy As Boolean = False, Optional pIsAy As Boolean = False)
Select Case pNmTypArg$
Case "$", "%", "!", "#", "&": ToStr_ArgDcl = pNmArg & pNmTypArg & IIf(pIsAy, "()", "")
Case Else:                    ToStr_ArgDcl = pNmArg & IIf(pIsAy, "()", "") & ":" & pNmTypArg
End Select
End Function
Function AppxToStr$(A As Excel.Application)
AppxToStr = "(" & A.Workbooks.Count & ") Wb. Wb1=" & WbToStr(A.Workbooks(1))
Exit Function
R: AppxToStr = "AppxToStr error.  Msg=" & Err.Description
End Function
Function ToStr_DPgmPrm$() ' (pDPgm As d_Pgm, pAyDPrm() As d_Arg)
If IsNothing(pDPgm) Then ToStr_DPgmPrm = "--Nothing--"
Dim mA$
With pDPgm
    Dim mNmRetTyp$, mNmAs
    Select Case .x_NmTypRet
    Case "#", "$", "%", "!", "&": mNmRetTyp = .x_NmTypRet
    Case Else: mNmAs = " As " & .x_NmTypRet
    End Select
    ToStr_DPgmPrm = Q_MrkUp(Fmt_Str("{0}{1}{2}{3}(){4}" _
        , IIf(.x_IsPrivate, "Private ", "") _
        , FctTyToStr(.x_TypFct) & " " _
        , Join_Lv(".", .x_NmPrj, .x_Nmm, .x_NmPrc) _
        , mNmRetTyp _
        , mNmAs), "DPgm") & vbLf & Q_MrkUp(.x_Aim, "Aim")
End With
End Function
Function ToStr_DPgm$(pDPgm As d_Pgm, Optional pByEle As Boolean = False)
If IsNothing(pDPgm) Then ToStr_DPgm = "--Nothing--"
Dim mA$
With pDPgm
    Dim mNmTypRet$, mNmAs
    Select Case .x_NmTypRet
    Case "#", "$", "%", "!", "&": mNmTypRet = .x_NmTypRet
    Case Else: mNmAs = " As " & .x_NmTypRet
    End Select
    ToStr_DPgm = Q_MrkUp(Fmt_Str("{0}{1}{2}{3}(){4}" _
        , IIf(.x_IsPrivate, "Private ", "") _
        , FctTyToStr(.x_TypFct) & " " _
        , Join_Lv(".", .x_NmPrj, .x_Nmm, .x_NmPrc) _
        , mNmTypRet _
        , mNmAs), "DPgm") & vbLf & Q_MrkUp(.x_Aim, "Aim")
End With
End Function
Function ToStr_DArg$() ' pDArg As d_Arg, Optional pByEle As Boolean = False)
If TypeName(pDArg) = "Nothing" Then ToStr_DArg = "Nothing": Exit Function
Dim mA$
With pDArg
    If pByEle Then
        ToStr_DArg = Fmt_Str("IsOpt={0}|IsByVal={1}|NmArg={2}|NmTypArg={3}|DftVal={4}" _
            , .x_IsOpt, .x_IsByVal, .x_NmArg, .x_NmTypArg, .x_DftVal)
        Exit Function
    End If
    If .x_IsOpt Then mA = "Optional "
    If .x_IsByVal Then mA = mA & ""
    mA = mA & .x_NmArg
    Select Case .x_NmTypArg
    Case "#", "$", "%", "&", "!":
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & .x_NmTypArg & "()"
        Else
            mA = mA & .x_NmTypArg
        End If
    Case "String"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "$()"
        Else
            mA = mA & "$"
        End If
    Case "Integer"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "%()"
        Else
            mA = mA & "%"
        End If
    Case "Single"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "!()"
        Else
            mA = mA & "!"
        End If
    Case "Long"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "&()"
        Else
            mA = mA & "&"
        End If
    Case "Double"
        If Right(.x_NmArg, 2) = "()" Then
            mA = Left(mA, Len(mA) - 2) & "#()"
        Else
            mA = mA & "#"
        End If
    Case Else: mA = mA & " As " & .x_NmTypArg
    End Select
    Select Case VarType(.x_DftVal)
    Case vbEmpty, vbNull
    Case vbString:
        If Trim(.x_DftVal) <> "" Then mA = mA & " = " & .x_DftVal
    Case Else: Stop
    End Select
End With
ToStr_DArg = mA
End Function
Function ToStr_AyDArg$() 'pAyDArg() As d_Arg, Optional pByEle As Boolean = False)
Dim N%: 'N = SzDArg(pAyDArg)
If N = 0 Then ToStr_AyDArg = "--NoArg--": Exit Function
Dim J%, mA$
For J = 0 To N - 1
    'mA = Add_Str(mA, ToStr_DArg(pAyDArg(J), pByEle), vbLf)
Next
ToStr_AyDArg = mA
End Function
Function ToStr_Md$(pMd As CodeModule)
On Error GoTo R
ToStr_Md = pMd.Parent.Collection.Parent.Name & "." & pMd.Name
Exit Function
R: ToStr_Md = "Err: ToStr_Md(pMd). Msg=" & Err.Description
End Function
Function ToStr_Prj$(pPrj As vbproject)
On Error GoTo R
ToStr_Prj = pPrj.Name
Exit Function
R: ToStr_Prj = "Err: ToStr_Prj(pPrj). Msg=" & Err.Description
End Function
Function ToStr_TypCmp$(pTypCmp As VBIDE.vbext_ComponentType)
Select Case pTypCmp
Case VBIDE.vbext_ComponentType.vbext_ct_ActiveXDesigner:    ToStr_TypCmp = "ActX"
Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule:        ToStr_TypCmp = "Class"
Case VBIDE.vbext_ComponentType.vbext_ct_Document:           ToStr_TypCmp = "Doc"
Case VBIDE.vbext_ComponentType.vbext_ct_MSForm:             ToStr_TypCmp = "Frm"
Case VBIDE.vbext_ComponentType.vbext_ct_StdModule:          ToStr_TypCmp = "Mod"
Case Else: ToStr_TypCmp = "Unknow(" & pTypCmp & ")"
End Select
End Function
Function ToStr_Vayv$(pVayv, Optional pQ$ = "", Optional pSepChr$ = CtCommaSpc)
If TypeName(pVayv) = "Nothing" Then ToStr_Vayv = "": Exit Function
Dim mAy(): mAy = pVayv
ToStr_Vayv = ToStr_AyV(mAy, pQ, pSepChr)
End Function

Function ToStr_Pkey$(pNmt$)
Const cSub$ = "ToStr_Pkey"
'Aim: Find the Pkey of given pNmt
On Error GoTo R
Dim I%
For I% = 0 To CurrentDb.TableDefs(pNmt).Indexes.Count - 1
    If CurrentDb.TableDefs(pNmt).Indexes(I).Primary Then
        Dim J%, mA$
        For J = 0 To CurrentDb.TableDefs(pNmt).Indexes(I).Fields.Count - 1
            mA = Add_Str(mA, CurrentDb.TableDefs(pNmt).Indexes(I).Fields(J).Name)
        Next
        ToStr_Pkey = mA
        Exit Function
    End If
Next
ss.A 1, "No Pkey index"
R: ss.R
E:
End Function

Function ToStr_Pkey__Tst()
Debug.Print ToStr_Pkey("mstAllBrand")
Debug.Print ToStr_Pkey("mstAllBrandaa")
End Function

Function ToStr_NmV$(pNm$, pV)
ToStr_NmV = pNm & "=[" & pV & "]"
End Function
Function ToStr_HostSts$(pHostSts As eHostSts)
Dim mA$
Select Case pHostSts
Case e1Rec: mA = "e1Rec"
Case e0Rec: mA = "e0Rec"
Case e2Rec: mA = "e2Rec"
Case eHostCpyToFrm: mA = "eHostCpyToFrm"
Case eUnExpectedErr: mA = "eUnExpectedErr"
Case Else: mA = "ToStr_HostSts: " & pHostSts
End Select
ToStr_HostSts = mA
End Function
Function ToStr_AyLng$(pAyLng&())
Dim J%, N%: N = Sz(pAyLng): If N = 0 Then Exit Function
Dim mS$: mS = pAyLng(0)
For J = 1 To N - 1
    mS = mS & ", " & pAyLng(J)
Next
ToStr_AyLng = mS
End Function
Function ToStr_An2V_New$(pAn2V() As tNm2V, Optional pSkipNull As Boolean = False)
Dim J%, mA$, V
For J = 0 To Siz_An2V(pAn2V) - 1
    V = pAn2V(J).NewV
    If pSkipNull Then If IsNull(V) Then GoTo Nxt
    mA = Add_Str(mA, Q_V(V))
Nxt:
Next
ToStr_An2V_New = mA
End Function
Function ToStr_An2V$(pAn2V() As tNm2V, Optional pSepChr$ = vbLf)
Dim J%, mA$
For J = 0 To Siz_An2V(pAn2V) - 1
    mA = Add_Str(mA, ToStr_Nm2V(pAn2V(J)), pSepChr)
Next
ToStr_An2V = mA
End Function
Function ToStr_Nm2V_Set(pNm2V As tNm2V) As Boolean
With pNm2V
    ToStr_Nm2V_Set = .Nm & "=" & Q_V(.NewV)
End With
End Function
Function ToStr_Nm2V$(pNm2V As tNm2V)
Dim mIsEq As Boolean
With pNm2V
    If IfEq_Nm2V(mIsEq, pNm2V) Then GoTo E
    Dim mA$
    If mIsEq Then
        mA$ = "<-NoChg"
    Else
        mA$ = "<-[" & Q_V(.OldV) & "]"
    End If
    ToStr_Nm2V = .Nm & "=[" & Q_V(.NewV) & "]" & mA
End With
Exit Function
E: ToStr_Nm2V = "Er IsEq_Nm2V"
End Function
Function ToStr_Am$(pAm() As tMap _
    , Optional pBrkChr$ = "=" _
    , Optional pQ1$ = "" _
    , Optional pQ2$ = "" _
    , Optional pSepChr$ = CtCommaSpc _
    )
Dim N%: N% = Siz_Am(pAm): If N% = 0 Then Exit Function
Dim J%, A$
For J = 0 To N - 1
    With pAm(J)
        If .F1 = "" Then
            If .F2 = "" Then
                A = Add_Str(A, "", pSepChr)
            Else
                A = Add_Str(A, Q_S(.F2, pQ2), pSepChr)
            End If
        Else
            If .F2 = "" Then
                A = Add_Str(A, Q_S(.F1, pQ1), pSepChr)
            Else
                A = Add_Str(A, Q_S(.F1, pQ1) & pBrkChr & Q_S(.F2, pQ2), pSepChr)
            End If
        End If
    End With
Next
ToStr_Am = A
End Function
Function ToStr_AmF1$(pAm() As tMap, Optional pQ$ = "", Optional pSepChr$ = CtCommaSpc)
Dim N%: N = Siz_Am(pAm): If N% = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    Dim A$: A = Add_Str(A, Q_S(pAm(J).F1, pQ), pSepChr)
Next
ToStr_AmF1 = A
End Function

Function ToStr_AmF1__Tst()
Const cSub$ = "ToStr_AmF1_Tst"
Const cLm$ = "aaa=xxx,bbb=yyy,1111"
Dim mAm() As tMap: mAm = Get_Am_ByLm(cLm)
Debug.Print "Input-----"
Debug.Print "cLm: "; cLm$
Debug.Print "Output----"
Debug.Print "ToStr_AmF1(cLm): "; ToStr_AmF1(mAm)
End Function

Function ToStr_AmF2$(pAm() As tMap, Optional pQ$ = "", Optional pSepChr$ = CtCommaSpc)
'Aim: list the F2 of {pAm} ToStr
Dim N%: N = Siz_Am(pAm): If N% = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    Dim A$: A = Add_Str(A, Q_S(pAm(J).F2, pQ), pSepChr)
Next
ToStr_AmF2 = A
End Function

Function ToStr_AmF2__Tst()
Const cSub$ = "ToStr_AmF2_Tst"
Const cLm$ = "aaa=xxx,bbb=yyy,1111"
Dim mAm() As tMap: mAm = Get_Am_ByLm(cLm)
Debug.Print "Input-----"
Debug.Print "cLm: "; cLm$
Debug.Print "Output----"
Debug.Print "ToStr_AmF2(cLm): "; ToStr_AmF2(mAm)
End Function

Function ToStr_AyNm2V$(pAyNm2V() As tNm2V)
Stop
End Function
Function ToStr_Ays$(pAys$(), Optional pQ$ = "", Optional pSepChr$ = CtCommaSpc)
Dim N%: N = Sz(pAys): If N% = 0 Then Exit Function
Dim A$: A = Q_S(pAys(0), pQ)
Dim J%
For J = 1 To N - 1
    A = A & pSepChr & Q_S(pAys(J), pQ)
Next
ToStr_Ays = A
End Function
Function ToStr_AyBool$(pAyBool() As Boolean, Optional pQ$ = "", Optional pSepChr$ = CtCommaSpc)
Dim N%: N = Sz(pAyBool): If N% = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    Dim A$: A = Add_Str(A, Q_S(CStr(pAyBool(J)), pQ), pSepChr)
Next
ToStr_AyBool = A
End Function
Function ToStr_AyByt$(pAyByt() As Byte, Optional pQ$ = "", Optional pSepChr$ = CtCommaSpc)
Dim N%: N = Sz(pAyByt): If N% = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    Dim A$: A = Add_Str(A, Q_S(CStr(pAyByt(J)), pQ), pSepChr)
Next
ToStr_AyByt = A
End Function
Function ToStr_AyV$(pAyV(), Optional pQ$ = "", Optional pSepChr$ = CtCommaSpc)
Dim N%: N = Sz(pAyV): If N% = 0 Then Exit Function
Dim A$, J%
For J = 1 To N - 1
    If (VarType(pAyV(J)) And vbArray) = 0 Then
        A$ = Add_Str(A$, Q_S(pAyV(J), pQ), pSepChr)
    Else
        Dim mX$: mX = "Array(" & Sz(pAyV(J)) & ")"
        A$ = Add_Str(A$, Q_S(mX, pQ), pSepChr)
    End If
Next
ToStr_AyV = A$
Exit Function
E: ToStr_AyV = "Err: ToStr_AyV(pAyV).  Msg=" & Err.Description
End Function
Function ToStr_Coll$(pColl As VBA.Collection, Optional pSepChr$ = CtComma)
If IsNothing(pColl) Then ToStr_Coll = "#Nothing#": Exit Function
Dim mV, mA$
For Each mV In pColl
    mA = Add_Str(mA, CStr(mV), pSepChr)
Next
ToStr_Coll = mA
End Function

Function ToStr_Ctl$(pCtl As Access.Control, Optional pWithTag As Boolean = False)
On Error GoTo R
If pWithTag Then
    If IsNothing(pCtl.Tag) Then
        ToStr_Ctl = pCtl.Name
    Else
        ToStr_Ctl = pCtl.Name & "(" & pCtl.Tag & ")"
    End If
Else
    ToStr_Ctl = pCtl.Name
End If
Exit Function
R: ToStr_Ctl = "Err: ToStr_Ctl(pCtl).  Msg=" & Err.Description
End Function
Function ToStr_Ctls$(pCtls As Access.Controls, Optional pWithTag As Boolean = False, Optional pSepChr$ = CtComma)
On Error GoTo R
Dim mS$, iCtl As Access.Control
For Each iCtl In pCtls
    mS = Add_Str(mS, ToStr_Ctl(iCtl, pWithTag), pSepChr)
Next
ToStr_Ctls = mS
Exit Function
R: ToStr_Ctls = "Err: ToStr_Ctls(pCtls).  Msg=" & Err.Description
End Function
Function ToStr_Rel$(pNmRel$, Optional pDb As database)
On Error GoTo R
Dim mDb As database: Set mDb = DbNz(pDb)
Dim mRel As DAO.Relation: Set mRel = mDb.Relations(pNmRel)
ToStr_Rel = "Rel(" & pNmRel & "):" & mRel.Table & ";" & mRel.ForeignTable & ";" & ToStr_Flds_Rel(mRel.Fields)
Exit Function
R: ToStr_Rel = "Err: ToStr_Rel(" & pNmRel & ").  Msg=" & Err.Description
End Function

Function ToStr_Rel__Tst()
Dim mDb As database: If Opn_Db_RW(mDb, "C:\Tmp\ProjMeta\Meta\MetaAll.Mdb") Then Stop
Debug.Print ToStr_Rel("AcptR10", mDb)
Shw_DbgWin
End Function

Function ToStr_Db$(pDb As database)
If IsNothing(pDb) Then ToStr_Db = "Nothing": Exit Function
On Error GoTo R
ToStr_Db = pDb.Name
Exit Function
R: ToStr_Db = "Err: ToStr_Db(pDb).  Msg=" & Err.Description
End Function
Function ToStr_FldVal$(pFld As DAO.Field)
On Error GoTo R
ToStr_FldVal = pFld.Value
Exit Function
R: ToStr_FldVal = "#" & Err.Description & "#"
End Function
Function ToStr_Tbl$(pTbl As DAO.TableDef)
Const cSub$ = "ToStr_Tbl"
On Error GoTo R
With pTbl
    ToStr_Tbl = .Name
End With
Exit Function
R: ToStr_Tbl = "Err: ToStr_Tbl(pTbl).  Msg=" & Err.Description
End Function
Function ToStr_Fld$(pFld As DAO.Field, Optional pInclTyp As Boolean = False, Optional pInclVal As Boolean = False)
Const cSub$ = "ToStr_Fld"
On Error GoTo R
With pFld
    If pInclTyp Then
        If pInclVal Then ToStr_Fld = .Name & ":" & ToStr_TypFld(pFld) & "=" & ToStr_FldVal(pFld): Exit Function
        ToStr_Fld = .Name & ":" & ToStr_TypFld(pFld)
        Exit Function
    End If
    If pInclVal Then ToStr_Fld = .Name & "=" & Nz(.ValidateOnSet, "Null"): Exit Function
    ToStr_Fld = .Name
End With
Exit Function
R: ToStr_Fld = "Err: ToStr_Fld(pFld).  Msg=" & Err.Description
End Function
Function ToStr_Fld_Rel$(pFld As DAO.Field)
Const cSub$ = "ToStr_Fld_Rel"
On Error GoTo R
With pFld
    If .Name = .ForeignName Then
        ToStr_Fld_Rel = .Name
    Else
        ToStr_Fld_Rel = .Name & "=" & .ForeignName
    End If
End With
Exit Function
R: ToStr_Fld_Rel = "Err: ToStr_Fld(pFld).  Msg=" & Err.Description
End Function
Function ToStr_Flds_Rel$(pFlds As DAO.Fields, Optional pSepChr$ = CtComma)
Const cSub$ = "ToStr_Flds_Rel"
On Error GoTo R
If pFlds.Count = 0 Then ToStr_Flds_Rel = "": Exit Function
Dim mA$, iFld As DAO.Field, J%
For J = 0 To pFlds.Count - 1
    mA = Add_Str(mA, ToStr_Fld_Rel(pFlds(J)), pSepChr)
Next
ToStr_Flds_Rel = mA
Exit Function
R: ToStr_Flds_Rel = "Err: ToStr_Flds(pFlds,pSepChr).  Msg=" & Err.Description
End Function
Function ToStr_Flds$(pFlds As DAO.Fields, Optional pInclTyp As Boolean = False, Optional pInclVal As Boolean = False, Optional pSepChr$ = CtComma, Optional pBeg As Byte = 0, Optional pEnd As Byte = 255)
Const cSub$ = "ToStr_Flds"
On Error GoTo R
If pFlds.Count = 0 Then ToStr_Flds = "": Exit Function
Dim mA$, iFld As DAO.Field, J%
For J = pBeg To Fct.MinByt(pFlds.Count - 1, CInt(pEnd))
    mA = Add_Str(mA, ToStr_Fld(pFlds(J), pInclTyp, pInclVal), pSepChr)
Next
ToStr_Flds = mA
Exit Function
R: ToStr_Flds = "Err: ToStr_Flds(pFlds,pInclTypFld,pSepChr,pBeg,pEnd).  Msg=" & Err.Description
End Function
Function ToStr_Flds__Tst()
Const cNmt$ = "mstBrand"
Dim mFlds As DAO.Fields: Set mFlds = CurrentDb.TableDefs(cNmt).Fields
Debug.Print ToStr_Flds(CurrentDb.TableDefs(cNmt).Fields, True, True)
Shw_DbgWin
End Function
Function ToStr_Fld_Dcl$(pFld As DAO.Field)
On Error GoTo R
ToStr_Fld_Dcl = pFld.Name & " " & Cv_Fld2Dcl(pFld)
Exit Function
R: ToStr_Fld_Dcl = "Err: ToStr_Fld_Dcl(pFld).  Msg=" & Err.Description
End Function
Function ToStr_Flds_Dcl$(pFlds As DAO.Fields, Optional pSepChr$ = CtComma)
Dim mA$
Dim iFld As DAO.Field: For Each iFld In pFlds
    mA = Add_Str(mA, ToStr_Fld_Dcl(iFld), pSepChr)
Next
ToStr_Flds_Dcl = mA
End Function
'Function ToStr_FmRecs(oS$, Sql$, Optional pSep$ = CtCommaSpc) As Boolean
'Const cSub$ = "ToStr_FmRecs"
'On Error GoTo R
'oS = ""
'With CurrentDb.OpenRecordset(Sql)
'    While Not .EOF
'        oS = Add_Str(oS, pSep)
'        .MoveNext
'    Wend
'    .Close
'End With
'Exit Function
'R: ss.R
'E: ToStr_FmRecs = True: ss.B cSub, cMod, "Sql", Sql
'End Function
Function ToStr_Frm$(pFrm As Access.Form)
On Error GoTo R
ToStr_Frm = pFrm.Name
Exit Function
R: ToStr_Frm = "Err ToStr_Frm(pFrm).  Msg=" & Err.Description
End Function
Function ToStr_FYNo$(pFyNo As Byte)
ToStr_FYNo = "FY" & Format(pFyNo, "00")
End Function
Function ToStr_Lang$(pLang As eLang)
Select Case pLang
Case eLang.eSC: ToStr_Lang = "SimpChinese"
Case eLang.eTC: ToStr_Lang = "TradChinese"
Case Else: ToStr_Lang = "English"
End Select
End Function
Function LpApToStr$(pSepChr$, pLp$, ParamArray pAp())
Dim mAm() As tMap: If Brk_LpVv2Am(mAm, pLp, CVar(pAp)) Then GoTo X
LpApToStr = ToStr_Am(mAm, , , "[]", pSepChr)
X:
End Function

Function LpApToStr__Tst()
Debug.Print LpApToStr(vbLf, "aa,bb,,C", 1, 2, , 1)
End Function

Function ToStr_Map$(pMap As tMap _
    , Optional pBrkChr$ = "=" _
    , Optional pQ1$ = "" _
    , Optional pQ2$ = "" _
        )
With pMap
    ToStr_Map = Q_S(.F1, pQ1) & pBrkChr & Q_S(.F2, pQ2)
End With
End Function
Function ToStr_Nmt$(pNmt$, Optional pInclTypFld As Boolean = False, Optional pSepChr$ = CtComma, Optional pBeg As Byte = 0, Optional pEnd As Byte = 255, Optional pDb As database)
Const cSub$ = "ToStr_Nmt"
On Error GoTo R
ToStr_Nmt = ToStr_Flds(DbNz(pDb).TableDefs(pNmt).Fields, pInclTypFld, , pSepChr, pBeg, pEnd)
Exit Function
R: ToStr_Nmt = "Err: ToStr_Nmt(" & pNmt & CtComma & ToStr_Db(pDb) & ").  Msg=" & Err.Description
End Function
Function ToStr_Nmt_Dcl$(pNmt$, Optional pSepChr$ = CtComma)
Const cSub$ = "ToStr_Nmt_Dcl"
On Error GoTo R
ToStr_Nmt_Dcl = ToStr_Flds_Dcl(CurrentDb.TableDefs(pNmt).Fields, pSepChr)
Exit Function
R: ss.R
    ToStr_Nmt_Dcl = "Err: ToStr_Nmt_Dcl(" & pNmt & ").  Msg=" & Err.Description
End Function
Function ToStr_Pc$(pPc As PivotCache)
If IsNothing(pPc) Then ToStr_Pc = "#Nothing#": Exit Function
Dim mCmdTxt$: mCmdTxt = "CmdTxt<Nil>"
Dim mCnnStr$: mCnnStr = "CnnStr<Nil>"
Dim mRfhNam$: mRfhNam = "RfhNam<Nil>"
Dim mPcIdx%
On Error Resume Next
With pPc
    mCmdTxt = .CtCommandText
    mCnnStr = .Connection
    mRfhNam = .RefreshName
    mPcIdx = .Index
End With
On Error GoTo 0
ToStr_Pc = ToStr_LpAp(CtComma, "CmdTxt,PcIdx,RfhNam,CnnStr", mCmdTxt, mPcIdx, mRfhNam, mCnnStr)
End Function
Function ToStr_Prp$(pPrp As DAO.Property)
On Error GoTo R
Dim mNm$: mNm = pPrp.Name
ToStr_Prp = mNm & "=[" & pPrp.Value & "]"
Exit Function
R: ss.R
    ToStr_Prp = "Err: ToStr_Prp(" & mNm & ").  Msg=" & Err.Description
End Function
Function ToStr_Prps$(pPrps As DAO.Properties, Optional pSepChr$ = " ")
Dim mA$, J As Byte
On Error GoTo R
For J = 0 To pPrps.Count - 1
    mA = Add_Str(mA, ToStr_Prp(pPrps(J)), pSepChr)
Next
ToStr_Prps = mA
Exit Function
R: ss.R
    ToStr_Prps = "Err: ToStr_Prps(pPrps,pSepChr).  Msg=" & Err.Description
End Function
Function ToStr_Prps__Tst()
Dim mPrps As DAO.Properties: Set mPrps = CurrentDb.QueryDefs("ODBCSQry").Properties
Debug.Print ToStr_Prps(mPrps, vbLf)
End Function

Function ToStr_Pt$(pPt As PivotTable)
If IsNothing(pPt) Then ToStr_Pt = "#Nothing#": Exit Function
Dim mCmdTxt$: mCmdTxt = "CmdTxt<Nil>"
Dim mCnnStr$: mCnnStr = "CnnStr<Nil>"
Dim mPcRfhNm$: mPcRfhNm = "PcRfhNm<Nil>"
Dim mPcIdx%
On Error Resume Next
With pPt
    mCmdTxt = .PivotCache.CtCommandText
    mCnnStr = .PivotCache.Connection
    mPcRfhNm = .PivotCache.RefreshName
    mPcIdx = .PivotCache.Index
End With
On Error GoTo 0
ToStr_Pt = ToStr_LpAp(CtComma, "CmdTxt,PcIdx,PtNam,PcRfhNm,CnnStr", mCmdTxt, mPcIdx, pPt.Name, mPcRfhNm, mCnnStr)
End Function
Function ToStr_Qt$(pQt As QueryTable)
If IsNothing(pQt) Then ToStr_Qt = "#Nothing#": Exit Function
Dim mCmdTxt$: mCmdTxt = "CmdTxt<Nil>"
Dim mCnnStr$: mCnnStr = "CnnStr<Nil>"
On Error Resume Next
With pQt
    mCmdTxt = .CtCommandText
    mCnnStr = .Connection
End With
On Error GoTo 0
ToStr_Qt = ToStr_LpAp(CtComma, "CmdTxt,QtNam,CnnStr", mCmdTxt, pQt.Name, mCnnStr)
End Function
Function ToStr_Rge$(Rg As Range)
On Error GoTo R
ToStr_Rge = Rg.Parent.Name & "!" & Rg.Address
Exit Function
R: ToStr_Rge = "Err: ToStr_Rge(Rg).  Msg=" & Err.Description
End Function
Function ToStr_RgeCno$(RgCno As tRgeCno)
With RgCno
    ToStr_RgeCno = "C" & .Fm & "-" & .To
End With
End Function
Function ToStr_Rs$(pRs As DAO.Recordset, Optional pRsNam$ = "", Optional pSepChr$ = CtComma)
Dim mRet$
mRet = "Rs value:" & IIf(pRsNam = "", "", "(RsNam=[" & pRsNam & "])")
On Error GoTo R
Dim iFld As DAO.Field: For Each iFld In pRs.Fields
    With iFld
        mRet = mRet & CtComma & ToStr_Fld(iFld)
    End With
Next
ToStr_Rs = mRet
Exit Function
R: ToStr_Rs = "Err: ToStr_Rs(pRs,pRsNam).  Msg=" & Err.Description
End Function
Function ToStr_Rs_NmFld$(pRs As DAO.Recordset, Optional pInclFldCnt As Boolean = False)
On Error GoTo R
Dim mRet$: If pInclFldCnt Then mRet = "NFld(" & pRs.Fields.Count & ") "
Dim iFld As DAO.Field
For Each iFld In pRs.Fields
    mRet = Add_Str(mRet, iFld.Name)
Next
ToStr_Rs_NmFld = mRet
Exit Function
R: ss.R
    ToStr_Rs_NmFld = "Err: ToStr_Rs_NmFld(pRs).  Msg=" & Err.Description
End Function
Function ToStr_Sq$(pSq As tSq)
With pSq
    ToStr_Sq = "(R" & .R1 & ",C" & .C1 & ") - (R" & .R2 & ",C" & .C2 & ")"
End With
End Function
Function ToStr_TypDta$(pTypDta As DAO.DataTypeEnum)
Select Case pTypDta
Case dbBigInt:  ToStr_TypDta = "BigInt": Exit Function
Case dbBinary:  ToStr_TypDta = "Binary": Exit Function
Case dbBoolean: ToStr_TypDta = "YesNo":   Exit Function
Case dbByte:    ToStr_TypDta = "Byte":   Exit Function
Case dbChar:    ToStr_TypDta = "Char": Exit Function
Case dbCurrency: ToStr_TypDta = "Currency": Exit Function
Case dbDate:    ToStr_TypDta = "Date": Exit Function
Case dbDecimal: ToStr_TypDta = "Decimal": Exit Function
Case dbDouble:  ToStr_TypDta = "Double": Exit Function
Case dbFloat:   ToStr_TypDta = "Float": Exit Function
Case dbGUID:    ToStr_TypDta = "GUID": Exit Function
Case dbInteger: ToStr_TypDta = "Int": Exit Function
Case dbLong:    ToStr_TypDta = "Long": Exit Function
Case dbLongBinary: ToStr_TypDta = "LongBinary": Exit Function
Case dbMemo:    ToStr_TypDta = "Memo": Exit Function
Case dbNumeric: ToStr_TypDta = "Numeric": Exit Function
Case dbSingle:  ToStr_TypDta = "Single": Exit Function
Case dbText:    ToStr_TypDta = "Text": Exit Function
Case dbTime:    ToStr_TypDta = "Time": Exit Function
Case dbTimeStamp: ToStr_TypDta = "TimeStamp":  Exit Function
Case dbVarBinary: ToStr_TypDta = "VarBinary":    Exit Function
Case Else:      ToStr_TypDta = "Unknow FieldTyp(" & pTypDta & ")"
End Select
End Function
Function ToStr_TypFld$(pFld As DAO.Field)
With pFld
    If .Type = dbText Then
        ToStr_TypFld = ToStr_TypDta(.Type) & .Size
    Else
        ToStr_TypFld = ToStr_TypDta(.Type)
    End If
End With
End Function
Function ToStr_TypMsg$(pTypMsg As eTypMsg)
Select Case pTypMsg
Case eTypMsg.ePrmErr:    ToStr_TypMsg = "PrmErr"
Case eTypMsg.eCritical:  ToStr_TypMsg = "Critical"
Case eTypMsg.eTrc:       ToStr_TypMsg = "Trace"
Case eTypMsg.eWarning:   ToStr_TypMsg = "Warning"
Case eTypMsg.eSeePrvMsg: ToStr_TypMsg = "SeePrvMsg"
Case eTypMsg.eException: ToStr_TypMsg = "Exception"
Case eTypMsg.eUsrInfo:   ToStr_TypMsg = "User Information"
Case eTypMsg.eRunTimErr: ToStr_TypMsg = "RunTimErr"
Case eTypMsg.eImpossibleReachHere: ToStr_TypMsg = "ImpossibleReachHere"
Case eTypMsg.eQuit: ToStr_TypMsg = "Application Quit"
Case Else: ToStr_TypMsg = "??(" & pTypMsg & ")"
End Select
'    ePrmErr = 1
'    eCritical = 2
'    eTrc = 3
'    eWarning = 4
'    eSeePrvMsg = 5
'    eException = 6
'    eUsrInfo = 7
'    eRunTimErr = 8
'    eImpossibleReachHere = 9
End Function
Function ToStr_TypObj$(pTypObj As AcObjectType)
Select Case pTypObj
Case AcObjectType.acForm:    ToStr_TypObj = "Forms":     Exit Function
Case AcObjectType.acQuery:   ToStr_TypObj = "Queries":   Exit Function
Case AcObjectType.acTable:   ToStr_TypObj = "Tables":    Exit Function
Case AcObjectType.acReport:  ToStr_TypObj = "Reports":   Exit Function
End Select
ToStr_TypObj = "AcObjectType(" & pTypObj & ")"
End Function
Function ToStr_TypQry$(pTypQry As DAO.QueryDefTypeEnum)
Select Case pTypQry
Case DAO.QueryDefTypeEnum.dbQAction:    ToStr_TypQry = "Action"
Case DAO.QueryDefTypeEnum.dbQAppend:    ToStr_TypQry = "Append"
Case DAO.QueryDefTypeEnum.dbQCompound:  ToStr_TypQry = "Compound"
Case DAO.QueryDefTypeEnum.dbQCrosstab:  ToStr_TypQry = "Crosstab"
Case DAO.QueryDefTypeEnum.dbQDDL:       ToStr_TypQry = "DDL"
Case DAO.QueryDefTypeEnum.dbQDelete:    ToStr_TypQry = "DDL"
Case DAO.QueryDefTypeEnum.dbQMakeTable: ToStr_TypQry = "MakeTable"
Case DAO.QueryDefTypeEnum.dbQProcedure: ToStr_TypQry = "Procedure"
Case DAO.QueryDefTypeEnum.dbQSelect:    ToStr_TypQry = "Select"
Case DAO.QueryDefTypeEnum.dbQSetOperation:  ToStr_TypQry = "SetOperation"   'Union
Case DAO.QueryDefTypeEnum.dbQSPTBulk:       ToStr_TypQry = "SPTBulk"
Case DAO.QueryDefTypeEnum.dbQSQLPassThrough: ToStr_TypQry = "SqlPassThrough"
Case DAO.QueryDefTypeEnum.dbQUpdate:        ToStr_TypQry = "Update"
Case Else: ToStr_TypQry = "Unknown(" & pTypQry & ")"
End Select
End Function
Function ToStr_TblAtr$(TblAtr&)
Dim mA$
If TblAtr And DAO.TableDefAttributeEnum.dbAttachedODBC Then mA = "ODBC"
If TblAtr And DAO.TableDefAttributeEnum.dbAttachedTable Then mA = Add_Str(mA, "Lnk", " ")
If TblAtr And DAO.TableDefAttributeEnum.dbHiddenObject Then mA = Add_Str(mA, "Hide", " ")
If TblAtr And DAO.TableDefAttributeEnum.dbSystemObject Then mA = Add_Str(mA, "Sys", " ")
ToStr_TblAtr = mA
End Function
Function ToStr_V$(pV As Variant)    ' pV is array of variant
Const cSub$ = "ToStr_V"
If VarType(pV) <> vbArray + vbVariant Then ss.A 1, "pV must be VarTyp of Array+Var", , "VarTyp of pV", VarType(pV): GoTo E
Dim mAyV(): mAyV = pV
Dim A$: A$ = mAyV(0)
Dim J As Byte
For J = 1 To UBound(mAyV)
    A$ = A$ & CtComma & mAyV(J)
Next
ToStr_V = A$
Exit Function
E:
    ToStr_V = "Err: ToStr_V(pV).  Msg=" & Err.Description
End Function
Function ToStr_VBPrj$(pVBPrj As vbproject)
On Error GoTo R
ToStr_VBPrj = pVBPrj.Name
Exit Function
R: ss.R
    ToStr_VBPrj = "Err: ToStr_VBPrj(pVbPrj).  Msg=" & Err.Description
End Function
Function WbToStr$(A As Workbook)
On Error GoTo R
WbToStr = A.FullName
Exit Function
R: WbToStr = "WbToStr error: Msg=" & Err.Description
End Function
Function ToStr_Wrd$(pWrd As Word.Document)
On Error GoTo R
ToStr_Wrd = pWrd.FullName
Exit Function
R: ss.R
    ToStr_Wrd = "Err: ToStr_Wrd(pWrd).  Msg=" & Err.Description
End Function
Function ToStr_Ws$(pWs As Worksheet, Optional pInclNmWb As Boolean = False)
On Error GoTo R
If pInclNmWb Then ToStr_Ws = "Wb=" & WbToStr(pWs.Parent) & ", Ws=" & pWs.Name: Exit Function
ToStr_Ws = "Ws=" & pWs.Name
Exit Function
R: ss.R
    ToStr_Ws = "Err: ToStr_Ws(pWs,pInclWb).  Msg=" & Err.Description
End Function
Function ToStr_YrWk$(pYr As Byte, pWk As Byte)
ToStr_YrWk = "Yr" & Format(pYr, "00") & "_Wk" & Format(pWk, "00")
End Function


