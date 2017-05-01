Attribute VB_Name = "nPpt_Appp"
Option Compare Database
Option Explicit

Function GetShape_ByStr(pPpt As Presentation, pS$) As PowerPoint.Shape
Dim iSlide As Slide
Dim iShape As PowerPoint.Shape
Dim iTxtFme As TextFrame
Dim iTxtRge As TextRange
For Each iSlide In pPpt.Slides
    For Each iShape In iSlide.Shapes
        If iShape.Type = msoTextBox Then
            Set iTxtRge = iShape.TextFrame.TextRange.Find(pS)
            If TypeName(iTxtRge) <> "Nothing" Then Set GetShape_ByStr = iShape: Exit Function
        End If
    Next
Next
End Function

Function Opn_Fp(pFp$) As Boolean
On Error GoTo R
If Not IsFfn(pFp, pSilient:=False) Then GoTo E
Dim mPpt As New PowerPoint.Application
mPpt.Visible = True
mPpt.Presentations.Open pFp, Untitled:=msoTrue
Exit Function
R: ss.R
E: Opn_Fp = True
End Function

Function PptOpn(Fppt$) As PowerPoint.Presentation
Dim Ppt As PowerPoint.Application
'Set Ppt = Appp
'Ppt.WindowState = WindowMinimized
'Set O = Ppt.Presentations.Open(pFp, Untitled:=msoTrue)
End Function

Function Sav_Ppt(pPpt As PowerPoint.Presentation) As Boolean
On Error GoTo E
pPpt.Save
Exit Function
E: Sav_Ppt = True
End Function

Function SavAs_Ppt(pPpt As PowerPoint.Presentation, pFp$) As Boolean
On Error GoTo E
pPpt.SaveAs pFp
Exit Function
E: SavAs_Ppt = True
End Function

Sub ShpRplShape_ByFx(Shp As PowerPoint.Shape, Fx$)
Dim Obj As Object: Set Obj = Shp.Parent
If TypeName(Obj) <> "Slide" Then Er "Shp.Parent is not Slide"
Dim mSlide As Slide: Set mSlide = Shp.Parent
mSlide.Select
Shp.Delete
mSlide.Shapes.AddOLEObject Left:=20, Top:=20, Width:=670, Height:=500, FileName:=Fx
End Sub
