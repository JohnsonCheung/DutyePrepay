Attribute VB_Name = "nAcs_Sts"
Option Compare Database
Option Explicit

Sub StsClr()
Application.SysCmd acSysCmdClearStatus
End Sub

Sub StsShw(Msg$)
SysCmd acSysCmdSetStatus, Msg
End Sub
