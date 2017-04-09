Attribute VB_Name = "nDao_nTxtSpec_TxtSpecTbl"
Option Compare Database
Option Explicit

Function TxtSpecTblCrt(Optional A As database) As Boolean
'Aim: Create 2 tables (MSysIMEXSpecs & MSysIMEXColumns) in {A}
Const cSub$ = "TxtSpecTblCrt"
If Not IsTbl("MSysIMEXSpecs", A) Then
    TblCrt_ByFldDclStr "MSysIMEXSpecs", _
        "SpecName Text 64" & _
        ", SpecId Auto" & _
        ", DateDelim Text 2" & _
        ", DateFourDigitYear YesNo" & _
        ", DateLeadingZeros YesNo" & _
        ", DecimalPoint Text 2" & _
        ", DateOrder Int" & _
        ", FieldSeparator Text 2" & _
        ", FileType Int" & _
        ", SpecType Byte" & _
        ", StartRow Long" & _
        ", TextDelim Text 2" & _
        ", TimeDelim Text 2" _
        , 1, 2, A
End If
If Not IsTbl("MSysIMEXColumns", A) Then
    TblCrt_ByFldDclStr "MSysIMEXColumns", _
        "SpecId Long" & _
        ", FieldName Text 64" & _
        ", Attributes Long" & _
        ", DataType Int" & _
        ", IndexType Byte" & _
        ", SkipColumn YesNo" & _
        ", Start Int" & _
        ", Width Int" _
        , 2, 2, A
    Dim mDb As database: Set mDb = DbNz(A)
    mDb.Execute "Create index Index1 on MSysIMEXColumns (SpecId,Start)"
End If
Exit Function
R: ss.R
E: TxtSpecTblCrt = True: ss.B cSub, cMod, "A", ToStr_Db(A)
'     Format of pLoFld is xxx Text 10,....
'     Note: xxx may be in xx^xx format.  ^ means for space
'       TEXT,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO
End Function

Function TxtSpecTblCrt__Tst()
Dim mDb As database: If Crt_Db(mDb, "c:\aa.mdb", True) Then Stop
If TxtSpecTblCrt(mDb) Then Stop
Stop
End Function

