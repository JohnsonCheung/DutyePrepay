Attribute VB_Name = "nIde_nMth_MthNm"
Option Compare Database
Option Explicit

Function MthNmCur$(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
Dim Kind As vbext_ProcKind
Dim Lin&, B&
Md.CodePane.GetSelection Lin, B, B, B
MthNmCur = Md.ProcOfLine(Lin, Kind)
End Function

Function MthNmEns$(MthNm$, Optional A As CodeModule)
Dim Ny$(): Ny = MdMthNy(A)
MthNmEns = NmNxt(MthNm, Ny)
End Function

Function MthNmMdNy(PubMthNm$, Optional A As vbproject) As String()
'Return {MdNy} which contains Public {MthNm}
Dim MdAy() As CodeModule: MdAy = PjMdAy(A)
Dim OMdAy() As CodeModule: OMdAy = AySel(MdAy, "MdHasMth_Pub", PubMthNm)
MthNmMdNy = ObjAyStrPrp(OMdAy, "Name")
End Function

Function MthNmNz$(MthNm$, A As CodeModule)
Dim O$
If MthNm = "" Then O = MthNmCur(A) Else O = MthNm
If O = "" Then Er "MthNmNz: No MthNmCur"
MthNmNz = O
End Function
