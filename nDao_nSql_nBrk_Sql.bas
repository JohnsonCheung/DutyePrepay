Attribute VB_Name = "nDao_nSql_nBrk_Sql"
Option Compare Database
Option Explicit

Sub SqFmt__Tst()
Dim I, Q As QueryDef
For Each I In QryAy
    Set Q = I
    Debug.Print SqlFmt(Q.Sql)
    Debug.Print "---------------------"
Next
End Sub

Function SqlFmt$(Sql$)
SqlFmt = JnCrLf(SqlKWPhraseAy(Sql))
End Function

Function SqlFmTn$(Sql$)
'Aim Find {OTn} from {Sql} by looking up the token after "From"
Dim P%: P = InStr(Sql, "From ")
If P = 0 Then Exit Function
Dim S$: S = RplCrLf(Mid(Sql, P + Len("From ")))
Dim Ay$(): Ay = Split(S, " "): If AyIsEmpty(Ay) Then GoTo E
Dim I
For Each I In Ay
    I = Trim(I)
    If I <> "" Then SqlFmTn = RmvPfxAll(I, "("): Exit Function
Next
E: Er "Given {Sql} does not have token following [From]", Sql
End Function

Sub SqlFmTn__Tst()
Dim Q As QueryDef
Dim I
For Each I In QryAy
    Set Q = I
    Debug.Print Q.Name, SqlFmTn(Q.Sql)
Next
End Sub

Function SqlKWPhraseAy(Sql$) As String()
SqlKWPhraseAy = BrkKWAy(Sql, ApSy("SELECT", "LEFT JOIN", "INNER JOIN", "RIGHT JOIN", "INSERT INTO", "FROM"))
End Function

Function SqlTny(Sql$) As String()
'^^
End Function
