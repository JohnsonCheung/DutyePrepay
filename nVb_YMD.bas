Attribute VB_Name = "nVb_YMD"
Option Compare Database
Option Explicit

Type YMD
    Y As Byte
    M As Byte
    D As Byte
End Type

Function YMDCur() As YMD
YMDCur = DteYMD(Date)
End Function

Function YMDToStr$(A As YMD)
With A
    YMDToStr = "20" & Format(.Y, "00") & "-" & Format(.M, "00") & "-" & Format(.D, "00")
End With
End Function
