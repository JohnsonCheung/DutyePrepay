Attribute VB_Name = "ZZ_xLst"
'Option Compare Text
'Option Explicit
'Option Base 0
'Const cMod$ = cLib & ".Lst"
'Function Lst_CmdTxt(pWb As Workbook, Optional pFno As Byte = 0) As Boolean
'Dim iWs As Worksheet, iQt As QueryTable, iPt As PivotTable
'For Each iWs In pWb.Worksheets
'    If iWs.PivotTables.Count > 0 Then
'        Prt_Ln pFno, Fct.UnderlineStr("Worksheet " & iWs.Name & " (PivotTables)", "-")
'        For Each iPt In iWs.PivotTables
'            Prt_Ln pFno, ToStr_Pt(iPt)
'        Next
'        Prt_Ln pFno
'    End If
'    If iWs.QueryTables.Count > 0 Then
'        Prt_Ln pFno, Fct.UnderlineStr("Worksheet " & iWs.Name & " (QueryTables)", "-")
'        For Each iQt In iWs.QueryTables
'            Prt_Ln pFno, ToStr_Qt(iQt)
'        Next
'        Prt_Ln pFno
'    End If
'Next
'End Function
'Function Lst_QryList(QryNmPfx$, Optional Sql_SubString$ = "") As Boolean
'Dim L%: L = Len(QryNmPfx)
'Dim iQry As QueryDef: For Each iQry In CurrentDb.QueryDefs
'    If Left(iQry.Name, L) = QryNmPfx Then If InStr(iQry.Sql, Sql_SubString) > 0 Then Debug.Print ToStr_TypQry(iQry.Type), iQry.Name
'Next
'End Function
'Function Lst_QryPrm_ByPfx(QryNmPfx$, Optional pFno As Byte = 0) As Boolean
'Dim L%: L = Len(QryNmPfx)
'Dim iQry As QueryDef: For Each iQry In CurrentDb.QueryDefs
'    If Left(iQry.Name, L) = QryNmPfx Then
'        If iQry.Parameters.Count > 0 Then
'            Prt_Str pFno, iQry.Name & "-----(Param)------>"
'            Dim iPrm As DAO.parameter
'            For Each iPrm In iQry.Parameters
'                Prt_Str pFno, iPrm.Name
'            Next
'            Prt_Ln pFno
'        End If
'    End If
'Next
'End Function
'
