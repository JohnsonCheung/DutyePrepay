Attribute VB_Name = "nIde_nPj_PjNm"
Option Compare Database
Option Explicit

Function PjNm_NewFalm$(PjNm, Optional Pth$)
PjNm_NewFalm = PjNm_NewFfn(PjNm, Pth, AppxExt)
End Function

Function PjNm_NewFfn$(PjNm, Pth$, Ext$)
PjNm_NewFfn = PjNzPth(Pth) & PjNm & Ext
End Function

Function PjNm_NewFmda$(PjNm, Optional Pth$)
PjNm_NewFmda = PjNm_NewFfn(PjNm, Pth, AppaExt)
End Function
