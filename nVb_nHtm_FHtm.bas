Attribute VB_Name = "nVb_nHtm_FHtm"
Option Compare Database
Option Explicit

Sub FhtmBrw(Fhtm$, Optional WinSty As VbAppWinStyle = vbMaximizedFocus)
Shell FmtQQ("""C:\Program Files (x86)\Google\Chrome\Application\Chrome.exe"" ""?""", Fhtm), WinSty
End Sub

Sub FhtmBrw__Tst()
Dim F$: F = TmpHtm
StrWrt "sdkfskdlfj", F
FhtmBrw F
End Sub
