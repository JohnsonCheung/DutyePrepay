Attribute VB_Name = "nVar_Var"
Option Compare Database
Option Explicit

Function IsNothing(V) As Boolean
IsNothing = TypeName(V) = "Nothing"
End Function

Sub VarAsg(A, OA)
If IsObject(A) Then
    Set OA = A
Else
    OA = A
End If
End Sub

Sub VarAsstEq(V1, V2)
ErAsst VarChkEq(V1, V2)
End Sub

Sub VarAsstEq__Tst()
VarAsstEq 1, 2
End Sub

Sub VarAsstSq(V, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst VarChkSq(V), Av
End Sub

Sub VarAsstSy(V, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst VarChkSy(V), Av
End Sub

Sub VarAsstSy__Tst()
VarAsstSy "sfsdf"

End Sub

Function VarBool(V) As Boolean
VarBool = V
End Function

Function VarChkEq(V1, V2) As Variant()

If VarType(V1) <> VarType(V2) Then
    Dim O()
    O = ErNew("Two V are dif VbTy", VarVbTyStr(V1), VarVbTyStr(V2))
    Push O, ErNew(".TypeName", TypeName(V1), TypeName(V2))
    VarChkEq = O
End If
If V1 = V2 Then Exit Function
VarChkEq = ErNew("Two V of same {VbTy} with Dif Val of {V1} and {V2}", VarVbTyStr(V1), V1, V2)
End Function

Function VarChkSq(V) As Variant()
If VarIsSq(V) Then Exit Function
VarChkSq = ErNew("Given V-{Type} is not a Sq", TypeName(V))
End Function

Function VarChkSy(V) As Variant()
If Not VarIsSy(V) Then
    VarChkSy = ErNew("Given V of {Ty} is not string array", TypeName(V))
End If
End Function

Function VarCv(V, VbTy As VbVarType)
Dim O
Select Case VbTy
Case VbVarType.vbBoolean:  O = CBool(V)
Case VbVarType.vbByte:     O = CByte(V)
Case VbVarType.vbCurrency: O = CCur(V)
Case VbVarType.vbDate:     O = CDate(V)
Case VbVarType.vbDecimal:  O = CDec(V)
Case VbVarType.vbDouble:   O = CDbl(V)
Case VbVarType.vbInteger:  O = CInt(V)
Case VbVarType.vbLong:     O = CLng(V)
Case VbVarType.vbLongLong: O = CLngLng(V)
Case VbVarType.vbSingle:   O = CSng(V)
Case VbVarType.vbString:   O = CStr(V)
Case Else: Er "Given {V} has {T} not in {Exp-VbTy}, cannot VarCv(V,T)", VarToStr(V), VbTyStr(VbTy), "YES BYT CUR DTE DEC DBL INT LNG LLG SNG STR"
End Select
VarCv = O
End Function

Function VarDaoTy(V) As DataTypeEnum
VarDaoTy = DaoTyNewByVbTy(VarType(V))
End Function

Function VarDte(V) As Date
VarDte = V
End Function

Function VarFmSemiColonFld(SemiColFld$, Optional Ty As VbVarType)

End Function

Function VarFmStr(Str$, Optional Ty As VbVarType)

End Function

Function VarInt%(V)
VarInt = V
End Function

Function VarIsBlank(V) As Boolean
Dim O As Boolean
If IsArray(V) Then VarIsBlank = Sz(V) = 0: Exit Function
Select Case VarType(V)
Case vbString: O = StrIsBlank(CStr(V))
Case vbNull, vbEmpty, vbError: O = True
End Select
VarIsBlank = O
End Function

Sub VarIsBlank__Tst()
VarIsBlank__Hlp
Dim V1()
Dim Act As Boolean
Act = VarIsBlank(V1)
Debug.Assert Act = True
End Sub

Function VarIsBool(V) As Boolean
VarIsBool = VarType(V) = vbBoolean
End Function

Function VarIsBoolTrue(V) As Boolean
If Not VarIsBool(V) Then Exit Function
VarIsBoolTrue = V
End Function

Function VarIsDbl(V) As Boolean
VarIsDbl = VarType(V) = vbDouble
End Function

Function VarIsDic(V) As Boolean
VarIsDic = TypeName(V) = "Dictionary"
End Function

Function VarIsEq(V1, V2, Optional IsExact As Boolean) As Boolean
If IsExact Then
    If VarType(V1) <> VarType(V2) Then Exit Function
End If
If IsNull(V1) Then Exit Function
If IsNull(V2) Then Exit Function
VarIsEq = V1 = V2
End Function

Function VarIsEq__Tst()
Dim mV1, mV2, mIsEq As Boolean
Dim mRslt As Boolean, mCase As Byte: mCase = 1
For mCase = 1 To 8
    Select Case mCase
    Case 1: mV1 = Null:    mV2 = Null
    Case 2: mV1 = Null:    mV2 = "1"
    Case 3: mV1 = "1":     mV2 = Null
    Case 4: mV1 = Null:    mV2 = 1
    Case 5: mV1 = 1:       mV2 = Null
    Case 6: mV1 = "1 ":    mV2 = "1"
    Case 7: mV1 = "1":     mV2 = "1 "
    Case 8: mV1 = 1:       mV2 = "1"
    End Select
    mRslt = VarIsEq(mV1, mV2)
    Debug.Print LpApToStr(vbTab & vbTab, "mRslt,mV1,mV2,mIsEq", mRslt, mV1 & "(" & TypeName(mV1) & ")", mV2 & "(" & TypeName(mV2) & ")", mIsEq)
Next
Debug.Assert VarIsEq(1, CByte(1)) = True
Debug.Assert VarIsEq(1, "1") = False
Debug.Assert VarIsEq(1, "1", True) = False

End Function

Function VarIsGe(V, Given) As Boolean
VarIsGe = V >= Given
End Function

Function VarIsGt(V, Given) As Boolean
VarIsGt = V > Given
End Function

Function VarIsInAp(V, ParamArray Ap()) As Boolean
Dim Av(), I
Av = Ap
For Each I In Av
    If V = I Then VarIsInAp = True: Exit Function
Next
End Function

Function VarIsPri(V) As Boolean
Select Case VarType(V)
Case vbBoolean, vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbLongLong, vbSingle, vbString: VarIsPri = True: Exit Function
Case Else
End Select
End Function

Function VarIsSq(V) As Boolean
If Not IsArray(V) Then Exit Function
Dim A&
On Error GoTo F
A = UBound(V, 2)
On Error GoTo T
A = UBound(V, 3)
F: Exit Function
T: VarIsSq = True
End Function

Sub VarIsSq__Tst()
Dim A()
Debug.Assert VarIsSq(A) = False
ReDim A(0)
Debug.Assert VarIsSq(A) = False
ReDim A(0, 0)
Debug.Assert VarIsSq(A) = True
ReDim A(0, 0, 0)
Debug.Assert VarIsSq(A) = False

'---
Dim B
Debug.Assert VarIsSq(B) = False
ReDim B(0)
Debug.Assert VarIsSq(B) = False
ReDim B(0, 0)
Debug.Assert VarIsSq(B) = True
ReDim B(0, 0, 0)
Debug.Assert VarIsSq(B) = False
End Sub

Function VarIsStr(V) As Boolean
VarIsStr = VarType(V) = vbString
End Function

Function VarIsSy(V) As Boolean
VarIsSy = VarType(V) = vbString + vbArray
End Function

Sub VarIsSy__Tst()
Dim A$()
Dim Ay
Dim B(): ReDim B(2)
Ay = A
Debug.Assert VarIsSy(A) = True
Debug.Assert VarIsSy(Ay) = True
Debug.Assert VarIsSy(B) = False
End Sub

Function VarLng&(V)
VarLng = V
End Function

Function VarOPtr(Obj) As LongPtr
VarOPtr = ObjPtr(Obj)
End Function

Function VarSemiColonFld$(V)
Dim A$: A = VarToStr(V)
If Not VarIsStr(V) Then VarSemiColonFld = A: Exit Function
VarSemiColonFld = EscTab(EscCR(EscLF(EsCtSemiColonColon(A))))
End Function

Function VarSemiColonFldRev(SemiColonFld$)
VarSemiColonFldRev = UnEscTab(UnEscCR(UnEscLF(UnEsCtSemiColonColon(SemiColonFld))))
End Function

Function VarStr&(V)
VarStr = V
End Function

Function VarToStr$(V)
If VarIsPri(V) Then VarToStr = V: Exit Function
If IsEmpty(V) Then Exit Function
If IsNull(V) Then Exit Function
VarToStr = "*" & TypeName(V)
End Function

Function VarVbTyStr$(V)
VarVbTyStr = VbTyStr(VarType(V))
End Function

Function VarVPtr(V) As LongPtr
VarVPtr = VarPtr(V)
End Function

Function VarWdt%(V)
VarWdt = Len(VarToStr(V))
End Function

Private Sub VarIsBlank__Hlp(Optional A)
Debug.Assert VarIsBlank(A) = True
End Sub
