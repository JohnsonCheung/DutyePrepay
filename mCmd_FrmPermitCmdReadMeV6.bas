Attribute VB_Name = "mCmd_FrmPermitCmdReadMeV6"
Option Compare Database
Option Explicit

Sub FrmPermitCmdReadMeV6()
DtBrw Dt
End Sub

Private Function Dt() As Dt
Dim O()
Push O, Array("Version", "6")
Push O, Array("Date", "2017-03-13")
Push O, Array("Enhancement", "Allow Import a permit from an Xlsx")
Push O, Array(".")
Push O, Array("Import Folder", "N:\SapAccessReports\DutyPrepay5\Import\")

Push O, Array("File Name", "<<PermitNo>>.xlsx")
Push O, Array("Columns in file", "Batch Number : TEXT")
Push O, Array("", "SKU          : NUMERIC")
Push O, Array("", "Order Qty#   : NUMERIC")
Push O, Array("", "The Xlsx must have above 3 columns of such TEXT or NUMERIC")
Push O, Array("")
Push O, Array("Import", "Put the Permit-Xlsx files to the Import Folder,")
Push O, Array("", "Run [Work with permit]")
Push O, Array("", "The program will find if any Permit-Xlsx, and,")
Push O, Array("", "create a permit record with [Blue] color under [Import-Button]")
Push O, Array("", "Locate the permit record with [Blue] color and click [Import]")
Push O, Array("", "The data from the Permit-Xlsx will be imported to the permit")
Push O, Array("", "Any data in the permit (if it is an old permit) will be overwritten by Permit-Xlsx")
Push O, Array("")
Push O, Array("Done", "The Permit-Xlsx file will be moved to [Done\YYYY-MM-DD hhmmss] under Import Folder")
Push O, Array("", "User can click [Edit] to edit the newly imported permit")
Push O, Array("")
Push O, Array("Delete", "Delete function is added to allow user to delete a permit")
Dt = DtNew(LvsSplit("Item Description"), O)
End Function


