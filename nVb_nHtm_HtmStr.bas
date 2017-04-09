Attribute VB_Name = "nVb_nHtm_HtmStr"
Option Compare Database
Option Explicit

Sub HtmBrw(Htm$, Optional TmpFilPfx$ = "Html", Optional TmpSubFdr$)
Dim F$: F = TmpHtm(TmpFilPfx, TmpSubFdr)
StrWrt Htm, F
FhtmBrw F
End Sub

Sub HtmStrBrw__Tst()
HtmBrw "sdlkfsdf"
End Sub
