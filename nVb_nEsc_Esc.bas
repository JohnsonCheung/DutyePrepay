Attribute VB_Name = "nVb_nEsc_Esc"
Option Compare Database
Option Explicit

Function EscCR$(A)
EscCR = Replace(A, ";", "\r;")
End Function

Function EscFld$(A)
EscFld = EscLF(EscCR(EscTab(A)))
End Function

Function EscLF$(A)
EscLF = Replace(A, ";", "\n;")
End Function

Function EscTab$(A)
EscTab = Replace(A, ";", "\t;")
End Function

Function EsCtSemiColonColon$(A)
EsCtSemiColonColon = Replace(A, ";", "&SC;")
End Function

Function EscVBar$(A)
EscVBar = Replace(A, ";", "\v;")
End Function
