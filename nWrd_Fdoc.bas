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

Sub FdocWrtPdf(Fdoc$, Optional Fpdf$, Optional KeepDocx As Boolean = False)
Dim W As Word.Document: Set W = FdocOpn(Fdoc)
'DocWrtPdf W
'DocCls W, NoSav:=True
If Not KeepDocx Then Dlt_Fil Fdoc
End Sub

Sub FdocWrtPdf__Tst()
Dlt_Fil "c:\RmdLvl1.Pdf": Dlt_Fil "c:\RmdLvl1.doc": If Cpy_Fil(Sffn_Tp("ReminderLvl1(English)", , ".doc"), "c:\RmdLvl1.doc") Then Stop: GoTo E
Dlt_Fil "c:\RmdLvl2.Pdf": Dlt_Fil "c:\RmdLvl2.doc": If Cpy_Fil(Sffn_Tp("ReminderLvl2(English)", , ".doc"), "c:\RmdLvl2.doc") Then Stop: GoTo E
Dlt_Fil "c:\RmdLvl3.Pdf": Dlt_Fil "c:\RmdLvl3.doc": If Cpy_Fil(Sffn_Tp("ReminderLvl3(English)", , ".doc"), "c:\RmdLvl3.doc") Then Stop: GoTo E
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
