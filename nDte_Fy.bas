Attribute VB_Name = "nDte_Fy"
Option Compare Database
Option Explicit

Function Fy$(pYYMM%)
'500 => FY06
If pYYMM = 9999 Then
    Fy = "FY07"
    Exit Function
End If
If pYYMM Mod 100 = 0 Then
    Fy = "FY" & VBA.Format((pYYMM \ 100) + 1, "00")
    Exit Function
End If
If pYYMM Mod 100 = 1 Then
    Fy = "FY" & VBA.Format(pYYMM \ 100, "00")
    Exit Function
End If
Fy = "FY" + VBA.Format((pYYMM \ 100) + 1, "00")
End Function

Function FyCur$()
Dim mYYMM%
mYYMM = VBA.Format(Date, "yymm")
FyCur = Fy(mYYMM)
End Function

Function FyPrev$(pNPrev_Year As Byte)
FyPrev = Fy((Year(Date) - pNPrev_Year - 2000) * 100 + Month(Date))
End Function

Function FyStartDate(pFy$) As Long
FyStartDate = "20" & VBA.Format(CInt(Right(pFy, 2)) - 1, "00") & "0201"
End Function
