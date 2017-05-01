Attribute VB_Name = "nDao_Flds"
Option Compare Database
Option Explicit

Function Flds(T, Optional A As database) As DAO.Fields
Set Flds = Tbl(T, A).Fields
End Function

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
FldsFny = OyPrp_Nm(FldsFldAy(Flds, FstNFld))
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

Function FldsToStr$(F As DAO.Fields, Optional InclTy As Boolean, Optional InclVal As Boolean)
Dim O$(), I, II As Field
For Each I In F
    Set II = I
    Push O, FldToStr(II, InclTy, InclVal)
Next
FldsToStr = Jn(O, vbCrLf)
End Function
