Attribute VB_Name = "nVb_SeqNm"
Option Compare Database
Option Explicit

Function SeqNmRbr(SeqNy$()) As String()
If AyIsEmpty(SeqNy) Then Exit Function
'---------------------------
Dim J&, I
'---------------------------
'Dim ASrtP1$()
'Dim ASrtP2$()
'Dim AIdx&()
'    ReDim P1(OU)
'    ReDim P2(OU)
'    J = 0
'    For Each I In SeqNy
'        With StrBrk1(I, "_")
'            P1(J) = .S1
'            P2(J) = .S2
'            J = J + 1
'        End With
'    Next
'    OIdx = AIdx
'    OSrtP1 = ASrtP1
'    OSrtP2 = ASrtP2
'Dim OP1$()  ' Part1 in original position
'Dim OP2$()  ' Part2 in original position
'---------------------------


'Dim Idx&(): Idx = AySrtIdx(SeqNy)
'Dim Srt$(): Srt = AySelByIdx(SeqNy, Idx)

Dim OSrtP1$() ' Part1 Sorted and renamed
Dim OSrtP2$() ' Part1 Sorted and renamed
Dim OIdx&() ' OIdx is binding OSrt & OP1 in this way OP1(i) is renamed as OSrt(OIdx(i))
Dim OU&
    OU = UB(SeqNy)
'---------------------------
Dim O$()
    ReDim O(OU)
    J = 0
    For Each I In OIdx
        O(J) = OSrtP1(I) & "_" & OSrtP2(I)
        J = J + 1
    Next
SeqNmRbr = O
End Function

