Attribute VB_Name = "nOfc_OfcCmd"
Option Compare Database
Option Explicit

Function OfcCmdClick() As Boolean
Dim BActCmdBarCtl As CommandBarControl:
    Set BActCmdBarCtl = Application.CommandBars.ActionControl
    If IsNothing(BActCmdBarCtl) Then Exit Function
    If BActCmdBarCtl.Type <> msoControlButton Then Exit Function

Dim AToolBarNm$
Dim CMd$
    AToolBarNm = BActCmdBarCtl.Parent.Name
    CMd = BActCmdBarCtl.Parameter
    'React on cmdBack
    Select Case CMd
    Case "cmdBack"
        If Application.CurrentObjectType = acQuery Then DoCmd.Close: Exit Function
        If Application.CurrentObjectType = acTable Then DoCmd.Close: Exit Function
        If Forms.Count = 1 Then Fct.Quit: Exit Function
        On Error Resume Next
        DoCmd.Close
        If Forms.Count = 1 Then Forms(1).SetFocus
        Exit Function
    End Select
Dim OFct$
    OFct = AToolBarNm & "_" & CMd
Run OFct    '
End Function
