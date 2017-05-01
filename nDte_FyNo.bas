Attribute VB_Name = "nDte_FyNo"
Option Compare Database
Option Explicit

Function FyNoToStr$(FyNo As Byte)
FyNoToStr = "FY" & Format(FyNo, "00")
End Function
