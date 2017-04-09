Attribute VB_Name = "nVb_Pdf"
Option Compare Text
Option Explicit

Sub PDFOpn(PDF$)
Shell FmtQQ("""C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"" ""{0}""", PDF), vbMaximizedFocus
End Sub

Sub PdfPrt(PDF$)
Shell FmtQQ("""C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"" /p ""{0}""", PDF), vbMaximizedFocus
End Sub
