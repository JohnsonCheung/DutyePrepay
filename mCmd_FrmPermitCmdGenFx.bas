Attribute VB_Name = "mCmd_FrmPermitCmdGenFx"
Option Compare Database
Option Explicit

Sub AA4()
FrmPermitCmdGenFx__Tst
End Sub

Sub FrmPermitCmdGenFx(PermitId&)
Dim oFx$
    oFx = ChqReqFx(PermitId)
    If FfnIsExist(oFx) Then
        If Not Start("Form exist, Regenerate?") Then
            FxWb(oFx).Application.Visible = True
            Exit Sub
        End If
    End If

Dim xTotCr@, xTot@, xPermitNo$, xPermitId$, xPostDate$, xPermitDate$, xBankCode$, xByUsr$
Dim xGLAc$, xGLAcName$
    With CurrentDb.OpenRecordset("Select * from Permit where Permit=" & PermitId)
        If .EOF Then .Close: If MsgBox("No records found in Table-Permit by PermitId=[" & PermitId & "]") = vbYes Then Stop Else Exit Sub
        xTot = !Tot
        xTotCr = -xTot
        xPermitNo = !PermitNo
        xPermitId = Format(!Permit, "00000")
        xPostDate = Format(!PostDate, "yyyy-mm-dd")
        xPermitDate = Format(!PermitDate, "yyyy-mm-dd")
        xBankCode = Nz(!BankCode.Value, "")
        xByUsr = Nz(!ByUsr.Value, "")
        xGLAc = Nz(!GLAc.Value, "")
        xGLAcName = Nz(!GLAcName.Value, "")
        .Close
    End With

Dim Fm$: Fm = FbCurPth & "Template\Template_DutyPrepay_Cheque_Request_Form.xls"
FfnCpy Fm, oFx, OvrWrt:=True

Dim xTxAmt@(), xBusArea$()
    Dim Sql$
    Sql = "SELECT [Business Area Code] as BusArea, Sum(x.Amt) as TxAmt" & _
    " FROM PermitD x" & _
    " Left JOIN qSKU s ON x.Sku = s.Sku" & _
    " WHERE x.Permit = " & PermitId & _
    " GROUP BY [Business Area Code];"
    Dim N%: N = 0
    With CurrentDb.OpenRecordset(Sql)
        If .EOF Then .Close: MsgBox "No record found in Table-PermitD by PermitId=[" & PermitId & "]": Exit Sub
        While Not .EOF
            ReDim Preserve xTxAmt(N), xBusArea(N)
            xTxAmt(N) = 0 - Nz(!TxAmt, 0)
            xBusArea(N) = Nz(!BusArea, "")
            N = N + 1
            .MoveNext
        Wend
        .Close
    End With

'' Fill in Ws by Variables

Dim oWb As Workbook
Dim oWs As Worksheet
    Set oWb = FxWb(oFx)
    Set oWs = oWb.Sheets(1)

Dim mRge As Range
Dim mCnoBusArea ' The column with {BusArea}
Dim mCnoTxAmt   ' The column with {TxAmt}
    Set mRge = oWb.Names("PrintArea").RefersToRange
    Dim mRnoBeg& ' The row with {BusArea}
    Dim iCell As Range
    For Each iCell In mRge
        Dim mV: mV = iCell.Value
        If VarType(mV) = vbString Then
            Dim mS$: mS = mV
            If Left(mS, 1) = "{" Then
                Select Case mS
                Case "{Tot}": iCell.Value = xTot
                Case "{TotCr}": iCell.Value = xTotCr
                Case "{PermitNo}": iCell.Value = xPermitNo
                Case "{PermitId}": iCell.Value = xPermitId
                Case "{PostDate}": iCell.Value = xPostDate
                Case "{PermitDate}": iCell.Value = xPermitDate
                Case "{BankCode}": iCell.Value = xBankCode
                Case "{ByUsr}": iCell.Value = xByUsr
                Case "{GLAc}": iCell.Value = xGLAc
                Case "{GLAcName}": iCell.Value = xGLAcName
                Case "{BusArea}": mRnoBeg = iCell.Row: mCnoBusArea = iCell.Column
                Case "{TxAmt}": mCnoTxAmt = iCell.Column
                Case "{TimeStamp}": iCell.Value = Format(Now, "yyyy-mm-dd hh:nn")
                End Select
            End If
        End If
    Next

'' Fill in Ws by TxAmt(), BusArea(), mRnoBeg, mCnoBusArea, mCnoTxAmt
If mRnoBeg = 0 Then
    MsgBox "No {BusArea} is found in the Template!!"
Else
    Dim J%
    Dim mRgeNxt As Range
    For J = 1 To UBound(xTxAmt)
        Set mRge = oWs.Rows(mRnoBeg)
        mRge.EntireRow.Select
        Selection.Copy
        Set mRgeNxt = oWs.Rows(mRnoBeg + 1)
        mRgeNxt.EntireRow.Select
        oWs.Paste
    Next
    For J = 0 To UBound(xTxAmt)
        Set mRge = oWs.Cells(J + mRnoBeg, mCnoTxAmt)
        mRge.Value = xTxAmt(J)
        Set mRge = oWs.Cells(J + mRnoBeg, mCnoBusArea)
        mRge.Value = xBusArea(J)
    Next
End If
SqlRun "SELECT x.Sku, qSKU.[SKU Description], x.Amt, x.Rate, x.Qty INTO [@Permit]" & _
" FROM Permit AS a INNER JOIN (PermitD AS x LEFT JOIN qSKU ON x.Sku = qSKU.Sku) ON a.Permit = x.Permit" & _
" WHERE x.Permit = " & PermitId & _
" ORDER BY x.SeqNo;"
WbRfh oWb
WbSav oWb
oWb.Application.Visible = True
End Sub

Sub FrmPermitCmdGenFx__Tst()
FrmPermitCmdGenFx 1692
End Sub

