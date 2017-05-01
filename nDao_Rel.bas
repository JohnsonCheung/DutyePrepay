Attribute VB_Name = "nDao_Rel"
Option Compare Database
Option Explicit

Function RelCrt(RelNm$, TFm$, TTo$, pLmFld$ _
    , Optional IsIntegral As Boolean, Optional IsCascadeUpd As Boolean, Optional IsCascadeDlt As Boolean, Optional A As database) As Boolean
'Aim: Create a relation. {pLmFld} is format of xx=yy,cc,dd=ee
'Dim Db As database: Set Db = DbNz(A)
'If DbHasRel(RelNm) Then Er "Given RelNm exist"
'Dim OAtr As DAO.RelationAttributeEnum
'If Not IsIntegral Then OAtr = dbRelationDontEnforce
'If IsCascadeUpd Then OAtr = OAtr Or dbRelationUpdateCascade
'If IsCascadeDlt Then OAtr = OAtr Or dbRelationDeleteCascade
'Dim O As DAO.Relation: Set O = Db.CreateRelation(RelNm, TFm, TTo, OAtr)
'Dim J%
'For J = 0 To Siz_Am(mAm) - 1
'    With mAm(J)
'        O.Fields.Append O.CreateField(.F1)
'        O.Fields(.F1).ForeignName = .F2
'    End With
'Next
'Db.Relations.Append O
End Function

Function RelCrt__Tst()
RelCrt "xxx#xx", "0Rec", "1Rec", "x", True, True, True
End Function

Sub RelDrp(RelNm$, Optional A As database)
DbNz(A).Relations.Delete RelNm
End Sub

Sub RelDrpAll(Optional A As database)
Dim Db As database: Set Db = DbNz(A)
With Db.Relations
    While .Count >= 1
        .Delete Db.Relations(0).Name
    Wend
End With
End Sub

Sub RelDrpAll__Tst()
Dim Db As database
Dim Fb$
If Opn_Db(Db, Fb, False) Then Stop
If Dlt_RelAll(Db) Then Stop
Db.Close
If Opn_CurDb(G.gAcs, Fb) Then Stop
G.gAcs.Visible = True
End Sub

Sub RelDrpIfExist(RelNm$, Optional A As database)
Dim Db As database: Set Db = DbNz(A)
If DbHasRel(RelNm, Db) Then RelDrp RelNm, Db
End Sub

Function RelToStr$(RelNm$, Optional A As database)
On Error GoTo R
Dim Db As database: Set Db = DbNz(A)
Dim mRel As DAO.Relation: Set mRel = A.Relations(RelNm)
'RelToStr = "Rel(" & RelNm & "):" & mRel.Table & ";" & mRel.ForeignTable & ";" & FldsToStr_Rel(mRel.Fields)
Exit Function
R: RelToStr = "Err: RelToStr(" & RelNm & ").  Msg=" & Err.Description
End Function

Function RelToStr__Tst()
Dim mDb As database: If Opn_Db_RW(mDb, "C:\Tmp\ProjMeta\Meta\MetaAll.Mdb") Then Stop
Debug.Print RelToStr("AcptR10", mDb)
End Function

