Attribute VB_Name = "nFs_CurPth"
Option Compare Database
Option Explicit

Private X_CurPth$()

Sub CurPthPop()
CurPthSet Pop(X_CurPth)
End Sub

Sub CurPthPush(NewCurPth$)
Push X_CurPth, CurDir
CurPthSet NewCurPth
End Sub

Sub CurPthSet(NewCurPth$)
ChDir NewCurPth
ChDrive Left(NewCurPth, 2)
End Sub
