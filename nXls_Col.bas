Attribute VB_Name = "nXls_Col"
Option Compare Database
Option Explicit

Function ColNxtN$(Col$, NCol%)
If NCol = 0 Then ColNxtN = Col: Exit Function
Dim A$
    A = UCase(Col)
If Len(A) = 1 Then
    If A = "Z" Then ColNxtN = "AA": Exit Function
    ColNxtN = Chr(Asc(Col) + 1)
    Exit Function
End If
If Len(Col) <> 2 Then Er "ColNxt: Given {Col} must be 1 or 2 char", Col
If Right(Col, 1) = "Z" Then ColNxtN = Chr(Asc(Left(Col, 1)) + 1) & "A": Exit Function
ColNxtN = Left(Col, 1) & Chr(Asc(Right(Col, 1)) + 1)
End Function

