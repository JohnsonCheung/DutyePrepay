Attribute VB_Name = "nXls_nFmt_ListObjFmtr"
Option Compare Database
Option Explicit
Private Type ZAlign
    Align() As XlHAlign
    AlignCno() As Integer
End Type
Private Type ZColr
    Colr() As Long
    Cno() As Integer
End Type
Private Type ZFormula
    Formula() As String
    Cno() As Integer
End Type
Private Type ZHdrColr
    Rno() As Integer
    Cno() As Integer
    Colr() As Long
End Type
Private Type ZVLin
    Cno() As Integer
End Type
Type ListObjFmtr
    HdrSq() As Variant
    
    VLinLeftCno() As Integer
    VLinRightCno() As Integer
    
    IsSepLin As Boolean
    SummaryCol As XlSummaryColumn
    SummaryRow As XlSummaryRow
    
    FormulaCno() As Integer
    Formula() As String
    
    FontColrCno() As Integer
    FontColr() As Long
    
    BackColrCno() As Integer
    BackColr() As Long
    
    SumTotColNm() As String
    SumAvgColNm() As String
    SumCntColNm As String
    
    LvlCno() As Integer
    LvlNo() As Integer
    
    HdrFontColrCno() As Integer
    HdrFontColrRno() As Integer
    HdrFontColr() As Long
    HdrBackColrCno() As Integer
    HdrBackColrRno() As Integer
    HdrBackColr() As Long
    
    AlignCno() As Integer
    Align() As XlHAlign
    
    NumFmtCno() As Integer
    NumFmt() As String
    
    Lbl() As String
    Fld() As String
End Type

Function ListObjFmtrNew(A As RgFmtDef) As ListObjFmtr
Dim J%
Dim O As ListObjFmtr
'Align
With ZAlign(A.Align)
    O.Align = .Align
    O.AlignCno = .AlignCno
End With

'BackColr
With ZColr(A.BackColr)
    O.BackColr = .Colr
    O.BackColrCno = .Cno
End With


'Fld
O.Fld = AyRmvAt(A.Fld)
    
'FontColr
With ZColr(A.FontColr)
    O.FontColrCno = .Cno
    O.FontColr = .Colr
End With

'Formula
With ZFormula(A.Formula)
    O.FormulaCno = .Cno
    O.Formula = .Formula
End With

'HdrBackColr
With ZHdrColr(A.HdrFontColrDrAy)
    O.HdrBackColr = .Colr
    O.HdrBackColrCno = .Cno
    O.HdrBackColrRno = .Rno
End With

'HdrFontColr
With ZHdrColr(A.HdrFontColrDrAy)
    O.HdrFontColr = .Colr
    O.HdrFontColrCno = .Cno
    O.HdrFontColrRno = .Rno
End With
X:
ListObjFmtrNew = O
End Function

Sub ZAlign__Tst()
Dim A()
Dim Act As ZAlign
    
A = Array("Align", "LR", "L", "R", "C", 1, Empty)
Act = ZAlign(A)

AyAsstEqExa Act.Align, Array(xlHAlignLeft, xlHAlignRight, xlHAlignCenter)
AyAsstEqExa Act.AlignCno, Array(2, 3, 4)
End Sub

Sub ZColr__Tst()
Dim A()
Dim Act As ZColr
    
A = Array("BackColr", 1232#, , 12321#, 12222#, Empty)
Act = ZColr(A)

AyAsstEqExa Act.Colr, Array(1232&, 12321&, 12222&)
AyAsstEqExa Act.Cno, Array(1, 3, 4)

End Sub

Sub ZHdrColr__Tst()
Dim A()
Dim Act As ZHdrColr
    
Push A, Array("BackColr", 1232#, , 12321#, 12222#, Empty)
Push A, Array("BackColr", 1232#, , 12321#, 12222#, Empty)
Push A, Array("BackColr", 1232#, , 12321#, 12222#, Empty)
Act = ZHdrColr(A)

AyAsstEqExa Act.Colr, Array(1232&, 12321&, 12222&, 1232&, 12321&, 12222&, 1232&, 12321&, 12222&)
AyAsstEqExa Act.Rno, Array(0, 0, 0, 1, 1, 1, 2, 2, 2)
AyAsstEqExa Act.Cno, Array(1, 3, 4, 1, 3, 4, 1, 3, 4)
End Sub

Private Function ZAlign(AlignDr()) As ZAlign
Dim O As ZAlign
Dim Align$, J%
For J = 1 To UB(AlignDr)
    If VarIsStr(AlignDr(J)) Then
        Align = UCase(AlignDr(J))
        If Len(Align) = 1 Then
            Select Case Align
            Case "L"
                Push O.AlignCno, J
                Push O.Align, XlHAlign.xlHAlignLeft
            Case "R"
                Push O.AlignCno, J
                Push O.Align, XlHAlign.xlHAlignRight
            Case "C"
                Push O.AlignCno, J
                Push O.Align, XlHAlign.xlHAlignCenter
            End Select
        End If
    End If
Next
ZAlign = O
End Function

Private Function ZColr(ColrDr()) As ZColr
Dim O As ZColr
Dim J%
For J = 1 To UB(ColrDr)
    If VarIsDbl(ColrDr(J)) Then
        Push O.Cno, J
        Push O.Colr, ColrDr(J)
    End If
Next
ZColr = O
End Function

Private Function ZFormula(FormulaDr()) As ZFormula
Dim J%, Formula$, O As ZFormula
For J = 1 To UB(FormulaDr)
    If VarIsStr(FormulaDr(J)) Then
        Formula = FormulaDr(J)
        Push O.Cno, J
        Push O.Formula, Formula
    End If
Next
ZFormula = O
End Function

Private Sub ZFormula__Tst()
Dim A()
Dim Act As ZFormula
    
A = Array("BackFormula", "sddsfsdf", , "lskdfsdlkf", "slkdfjsdf", Empty)
Act = ZFormula(A)

AyAsstEqExa Act.Formula, Array("sddsfsdf", "lskdfsdlkf", "slkdfjsdf")
AyAsstEqExa Act.Cno, Array(1, 3, 4)
End Sub

Private Function ZHdrColr(HdrColrDrAy()) As ZHdrColr
Dim O As ZHdrColr
Dim R&, Dr, J%
If Not AyIsEmpty(HdrColrDrAy) Then
    R = -1
    For Each Dr In HdrColrDrAy
        R = R + 1
        For J = 1 To UB(Dr)
            If VarIsDbl(Dr(J)) Then
                Push O.Cno, J
                Push O.Rno, R
                Push O.Colr, Dr(J)
                
            End If
        Next
        R = R + 1
    Next
End If
ZHdrColr = O
End Function
