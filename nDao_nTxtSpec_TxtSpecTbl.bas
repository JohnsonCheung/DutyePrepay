Attribute VB_Name = "nDao_nTxtSpec_TxtSpecTbl"
Option Compare Database
Option Explicit

Sub TxtSpecTblCrt(Optional A As database)
'Aim: Create 2 tables (MSysIMEXSpecs & MSysIMEXColumns) in {A}
Dim Db As database: Set Db = DbNz(A)
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
Db.Execute "Create index Index1 on MSysIMEXColumns (SpecId,Start)"
End Sub

Sub TxtSpecTblEns(Optional A As database)
If TxtSpecTblIsExist(A) Then Exit Sub
TxtSpecTblCrt A
End Sub

Function TxtSpecTblEns__Tst()
TxtSpecTblEns
End Function

Function TxtSpecTblIsExist(Optional A As database) As Boolean

End Function

