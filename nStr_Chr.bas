Attribute VB_Name = "nStr_Chr"
Option Compare Database
Option Explicit

Function ChrIsCap(S) As Boolean
Dim C%: C = FstAsc(S)
ChrIsCap = (65 <= C And C <= 90)
End Function

Function ChrIsDig(S$) As Boolean
Dim C$: C = FstChr(S)
ChrIsDig = ("0" <= C And C <= "9")
End Function

Function ChrIsLetter(S) As Boolean
Dim C$: C = UCase(FstChr(S))
If "A" <= C And C <= "Z" Then ChrIsLetter = True: Exit Function
End Function

Function ChrIsNmChr(S) As Boolean
Dim C$: C = UCase(FstChr(S))
ChrIsNmChr = True
If "A" <= C And C <= "Z" Then Exit Function
If "0" <= C And C <= "9" Then Exit Function
If C = "_" Then Exit Function
ChrIsNmChr = False
End Function

Function ChrIsPun(S) As Boolean
Const PunChrList$ = "~!@#$%^&*()-+={}[]:;'""<>,.?/"""
Dim C$: C = FstChr(S)
ChrIsPun = InStr(PunChrList, C) > 0
End Function

Sub ChrIsPun__Tst()
Debug.Assert ChrIsPun(".") = True
End Sub
