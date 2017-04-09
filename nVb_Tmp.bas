Attribute VB_Name = "nVb_Tmp"
Option Compare Database
Option Explicit
Public Fso As New FileSystemObject

Function TmpFb$(Optional Pfx$, Optional SubFdr$ = "")
TmpFb = TmpFil(".accdb", Pfx, SubFdr)
End Function

Function TmpFil$(Optional Ext$, Optional Fnn$ = "", Optional SubFdr$ = "")
TmpFil = TmpPth(SubFdr) & TimStmp(Fnn) & Ext
End Function

Function TmpFt$(Optional Pfx$ = "", Optional SubFdr$ = "")
TmpFt = TmpFil(".txt", Pfx, SubFdr)
End Function

Function TmpFx$(Optional FxFnn$, Optional SubFdr$ = "")
TmpFx = TmpFil(".xlsx", FxFnn, SubFdr)
End Function

Function TmpHtm$(Optional Pfx$ = "", Optional SubFdr$ = "")
TmpHtm = TmpFil(".html", Pfx, SubFdr)
End Function

Function TmpNm$(Optional Pfx$ = "N")
Static I&
Dim A$
If Pfx <> "" Then A = Pfx & "_"
TmpNm = A & Format(Now(), "YYYY_MM_DD_HHMMSS_") & I
I = I + 1
End Function

