Attribute VB_Name = "nIde_nMth_MthBrk"
Option Compare Database
Option Explicit
Type MthBrk
    mFY As String
    Ty As String
    PrpTy As String
    Nm As String
    RetTyChr As String
    PrmStr As String
    RetAs As String
End Type

Function MthBrkAy(Optional A As CodeModule) As MthBrk()
Dim B$(): B = MdBdyLy(A)
Dim O() As MthBrk, M As MthBrk
Dim U&, J&
For J = 0 To UB(B)
    If Not SrcLinIsMth(B(J)) Then GoTo Nxt
    M = MthBrkNew(B(J))
    ReDim Preserve O(U)
    O(U) = M
    U = U + 1
Nxt:
Next
MthBrkAy = O
End Function

Function MthBrkDr(A As MthBrk) As Variant()
Dim O()
With A
    Push O, .Nm
    Push O, .mFY
    Push O, .Ty
    Push O, .PrpTy
    Push O, .RetTyChr
    Push O, .RetAs
    Push O, .PrmStr
End With
MthBrkDr = O
End Function

Function MthBrkFny() As String()
MthBrkFny = Split("Nm Mdy Ty PrpTy RetTyChr RetTy PrmStr")
End Function

Function MthBrkIsEmpty(A As MthBrk) As Boolean
MthBrkIsEmpty = A.Nm = ""
End Function

Function MthBrkIsEmptyAy(A() As MthBrk) As Boolean
On Error GoTo X
Dim U&: U = UBound(A)
Exit Function
X:
MthBrkIsEmptyAy = True
End Function

Function MthBrkIsRetObj(A As MthBrk) As Boolean
With A
    If .RetTyChr <> "" Then Exit Function
    If IsSfx(.RetAs, "()") Then Exit Function
End With
MthBrkIsRetObj = True
End Function

Function MthBrkIsTth(A As MthBrk) As Boolean
If Not NmIsTstNm(A.Nm) Then Exit Function
If Not MthBrkIsSubNoPrm(A) Then
    Er "Given {MthBrk} is Tst-Nm, but it is not [Sub XXX()]", MthBrkToStr(A)
End If
MthBrkIsTth = True
End Function

Function MthBrkIsTth_Pfx(A As MthBrk) As Boolean
If Not IsPfx(A.Nm, "Tst_") Then Exit Function
If Not MthBrkIsSubNoPrm(A) Then Exit Function
MthBrkIsTth_Pfx = True
End Function

Function MthBrkIsTth_Pri(A As MthBrk) As Boolean
If Not MthBrkIsTth(A) Then Exit Function
MthBrkIsTth_Pri = A.mFY = "Private"
End Function

Function MthBrkIsTth_PriPfx(A As MthBrk) As Boolean
If Not MthBrkIsTth_Pfx(A) Then Exit Function
MthBrkIsTth_PriPfx = (A.mFY = "Prilic" Or A.mFY = "")
End Function

Function MthBrkIsTth_PriSfx(A As MthBrk) As Boolean
If Not MthBrkIsTth_Sfx(A) Then Exit Function
MthBrkIsTth_PriSfx = (A.mFY = "Prilic" Or A.mFY = "")
End Function

Function MthBrkIsTth_Pub(A As MthBrk) As Boolean
If Not MthBrkIsTth(A) Then Exit Function
MthBrkIsTth_Pub = (A.mFY = "Public" Or A.mFY = "")
End Function

Function MthBrkIsTth_PubPfx(A As MthBrk) As Boolean
If Not MthBrkIsTth_Pfx(A) Then Exit Function
MthBrkIsTth_PubPfx = (A.mFY = "Public" Or A.mFY = "")
End Function

Function MthBrkIsTth_PubSfx(A As MthBrk) As Boolean
If Not MthBrkIsTth_Sfx(A) Then Exit Function
MthBrkIsTth_PubSfx = (A.mFY = "Public" Or A.mFY = "")
End Function

Function MthBrkIsTth_Sfx(A As MthBrk) As Boolean
If Not IsSfx(A.Nm, "_Tst") Then Exit Function
If Not MthBrkIsSubNoPrm(A) Then Exit Function
MthBrkIsTth_Sfx = True
End Function

Function MthBrkMatch(A As MthBrk, MthNm$, PrpTy$) As Boolean
If MthNm = "" Then
    MthBrkMatch = True
    Exit Function
End If
If MthNm <> A.Nm Then Exit Function
If PrpTy = "" Then MthBrkMatch = True: Exit Function
MthBrkMatch = A.PrpTy = PrpTy
End Function

Function MthBrkNew(MthLin) As MthBrk
Dim A$: A = MthLin
Dim B$
Dim O As MthBrk
With O
    .mFY = ParseSy(A, ApSy("Public", "Private", "Friend")):   A = LTrim(A)
    .Ty = ParseSy(A, ApSy("Function", "Sub", "Property")): A = LTrim(A): If .Ty = "" Then Er "{MthLin} should one of [Function | Sub | Property]", MthLin
    If .Ty = "Property" Then
        .PrpTy = ParseSy(A, ApSy("Get", "Set", "Let")): A = LTrim(A): If .PrpTy = "" Then Er "{MthLin} should one of [Get | Set | Let] after [Property]", MthLin
    End If
    .Nm = ParseNm(A)
    .RetTyChr = ParseSy(A, ApSy("@", "!", "#", "$", "%", "&"))
    B = ParseStr(A, "("): If B <> "(" Then Er "[(] is missing in {MthLines}", MthLin
    .PrmStr = ParseTillClsBkt(A)
    B = ParseStr(A, " As ")
    .RetAs = A
End With
MthBrkNew = O
End Function

Function MthBrkNewMthNm(MthNm$, Optional A As CodeModule) As MthBrk
Dim Lin$: Lin = MthLin(MthNm, A)
If Lin = "" Then Exit Function
MthBrkNewMthNm = MthBrkNew(Lin)
End Function

Function MthBrkToStr$(A As MthBrk)
With A
Dim M$: M = AddSpcAft(.mFY)
Dim T$: T = AddSpcAft(.Ty)
Dim P$: P = AddSpcAft(.PrpTy)
Dim N$: N = .Nm
Dim C$: C = .RetTyChr
Dim pS$: pS = QuoteBkt(.PrmStr)
Dim R$: If .RetAs <> "" Then R = " AS " & .RetAs
End With
MthBrkToStr = M & T & P & N & C & pS & R
End Function

Sub MthBrkToStr__Tst()
Dim A As CodeModule: Set A = Md("mMthStru")
Dim B$(): B = MdBdyLy(A)
Dim J%
For J = 0 To UB(B)
    If SrcLinIsMth(B(J)) Then
        Debug.Assert MthBrkToStr(MthBrkNew(B(J))) = B(J)
    End If
Next
End Sub

Private Function MthBrkIsSubNoPrm(A As MthBrk) As Boolean
With A
    If .Ty <> "Sub" Then Exit Function
    If .PrmStr <> "" Then Exit Function
End With
MthBrkIsSubNoPrm = True
End Function
