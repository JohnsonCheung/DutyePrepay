Attribute VB_Name = "nIde_nPj_Fmda"
Option Compare Database
Option Explicit

Function FmdaCrt(Fmda$, Optional A As Access.Application) As vbproject
FfnAsstExt Fmda, ".mda", "FmdaCrt"
Dim App As Access.Application: Set App = AppaNz(A)
App.Visible = True
App.DBEngine.CreateDatabase Fmda, dbLangGeneral
App.OpenCurrentDatabase Fmda
Set FmdaCrt = App.Vbe.VBProjects(1)
End Function
