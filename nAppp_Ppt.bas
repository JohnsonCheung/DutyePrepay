Attribute VB_Name = "nAppp_Ppt"
Option Compare Database
Option Explicit

Function PptNew(Optional Fppt) As Presentation
Dim O As Presentation
'Set O = Appp.Presentations.Add
If Fppt <> "" Then O.SaveAs Fppt
End Function

