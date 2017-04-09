Attribute VB_Name = "nDao_Flds"
Option Compare Database
Option Explicit

Function FldsDr(A As DAO.Fields, Optional FstNFld% = 0) As Variant()
Dim N%:
    If FstNFld <= 0 Then
        N = A.Count
    Else
        N = FstNFld
    End If
Dim O()
    ReDim O(N - 1)
    Dim J%
    For J = 0 To N - 1
        O(J) = A(J).Value
    Next
FldsDr = O
End Function

Function FldsFldAy(A As DAO.Fields, Optional FstNFld% = 0) As DAO.Field()
Dim N%:
    If FstNFld <= 0 Then
        N = A.Count
    Else
        N = FstNFld
    End If
Dim O() As DAO.Field
    ReDim O(N - 1)
    Dim J%
    For J = 0 To N - 1
        Set O(J) = A(J)
    Next
FldsFldAy = O
End Function

Function FldsFny(Flds As DAO.Fields, Optional FstNFld%) As String()
FldsFny = ObjAyPrp(FldsFldAy(Flds, FstNFld), "Name", ApSy)
End Function

Sub FldsFny__Tst()
AyBrw FldsFny(CurrentDb.TableDefs("Permit").Fields, 2)
End Sub

Function FldsHasFld(Flds As DAO.Fields, F) As Boolean
Dim I As Field
For Each I In Flds
    If I.Name = F Then FldsHasFld = True: Exit Function
Next
End Function
