Attribute VB_Name = "nIde_nTok_MthLy"
Option Compare Database
Option Explicit

Function MthLyDefTokNy(SrcLy$()) As String()
'MthLy to Defined TokNy
Dim Fn$
Dim PrmNy$()
Dim DimNy$()
    Dim MthLin$: MthLin = SrcLyOneContinueLin(SrcLy, 0)
    Dim Brk As MthBrk: Brk = MthBrkNew(MthLin)
    Fn = Brk.Nm
    PrmNy = PrmStrToNy(Brk.PrmStr)
    DimNy = SrcLyDimNy(SrcLy)
MthLyDefTokNy = ApSy(Fn, PrmNy, DimNy)
End Function
