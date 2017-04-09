Attribute VB_Name = "nIde_nPj_PjRf"
Option Compare Database
Option Explicit

Function PjPjRfNewDrAy(A As vbproject) As Variant()
Dim O()
Dim I As VBIDE.Reference
For Each I In PjNz(A).References
    Push O, PjRfDr(I)
Next
PjPjRfNewDrAy = O
End Function

Sub PjRfAddPj(TarPjFfn$, Optional P As vbproject)
Dim Pj As vbproject: Set Pj = PjNz(P)
If PjRfIsExist(TarPjFfn, Pj) Then Exit Sub
Pj.References.AddFromFile TarPjFfn
End Sub

Sub PjRfBrw(Optional A As vbproject)
DtBrw PjRfDt1(A)
End Sub

Function PjRfDr(A As VBIDE.Reference) As Variant()
Dim O()
With A
    Push O, .Name
    Push O, .IsBroken
    Push O, .Description
    Push O, .FullPath
    Push O, .Major
    Push O, .Minor
    Push O, .Type
End With
PjRfDr = O
End Function

Function PjRfDt1(Optional P As vbproject) As Dt
PjRfDt1 = DtNew(LvsSplit("Name IsBroken Description FullPath Major Minor Type"), PjPjRfNewDrAy(P))
End Function

Function PjRfIsExist(TarPjFfn$, Optional P As vbproject) As Boolean
Dim T1$: T1 = FfnFnn(TarPjFfn)
Dim I As VBIDE.Reference
For Each I In P.References
    If Not I.IsBroken Then
        If I.Name = T1 Then PjRfIsExist = True: Exit Function
    End If
Next
End Function

Sub PjRfSetAutoRef()
Dim mDirCur$: mDirCur = CurrentDb.Name
Dim mP%: mP = InStrRev(mDirCur, "\")
mDirCur = Left(mDirCur, mP)
Dim mDirObj$: mDirObj = mDirCur & "Working\PgmObj\"
Dim mFfnModU$: mFfnModU = mDirObj & "mda"
If VBA.Dir(mFfnModU$) Then MsgBox (mFfnModU$ & "  not found."): Application.Quit
Dim iRef As Reference: For Each iRef In Application.References
    If iRef.Name = cLib Then Application.References.Remove iRef
Next
Application.References.AddFromFile mFfnModU
End Sub
