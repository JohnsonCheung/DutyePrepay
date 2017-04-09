Attribute VB_Name = "nVb_nEsc_UnEsc"
Option Compare Database
Option Explicit

Function UnEscCR$(A)
UnEscCR = Replace(A, "\r", ";")
End Function

Function UnEscFld$(A)
UnEscFld = UnEscLF(UnEscCR(UnEscTab(A)))
End Function

Function UnEscLF$(A)
UnEscLF = Replace(A, "\n", ";")
End Function

Function UnEscTab$(A)
UnEscTab = Replace(A, "\t", ";")
End Function

Function UnEsCtSemiColonColon$(A)
UnEsCtSemiColonColon = Replace(A, "&SC;", ";")
End Function

Function UnEscVBar$(A)
UnEscVBar = Replace(A, "\v", ";")
End Function
