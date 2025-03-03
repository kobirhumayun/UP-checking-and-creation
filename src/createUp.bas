Attribute VB_Name = "createUp"
Option Explicit


Private Function dealWithUpClause6(upClause6RangObj As Range, sourceDataAsDicUpIssuingStatus As Object) As Variant
    'this function put all the information clause 6 on new UP

        'keep only one buyer
        If upClause6RangObj.Rows.Count > 1 Then

            upClause6RangObj.Rows("2:" & upClause6RangObj.Rows.Count).EntireRow.Delete

        End If

        'insert rows as lc count, note already one row exist
        If sourceDataAsDicUpIssuingStatus.Count > 1 Then

            Dim i As Long
            For i = 1 To sourceDataAsDicUpIssuingStatus.Count - 1
                upClause6RangObj.Rows("2").EntireRow.Insert
            Next i

        End If

        Set upClause6RangObj = upClause6RangObj.Resize(sourceDataAsDicUpIssuingStatus.Count)

        If sourceDataAsDicUpIssuingStatus.Count = 1 Then

            upClause6RangObj(1, 14).value = sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(0))("NameofBuyers")

        Else

            Dim j As Long
            For j = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

                upClause6RangObj(j + 1, 14).value = j + 1 & ") " & sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("NameofBuyers")
                upClause6RangObj.Range("n" & j + 1 & ":z" & j + 1).Merge
            Next j

        End If

        upClause6RangObj.Range("b1:b" & upClause6RangObj.Rows.Count).Merge
        upClause6RangObj.Range("c1:m" & upClause6RangObj.Rows.Count).Merge

        Application.Run "utility_formating_fun.SetBorderInsideHairlineAroundThin", upClause6RangObj.Range("b1:z" & upClause6RangObj.Rows.Count)
        Application.Run "utility_formating_fun.setBorder", upClause6RangObj.Range("b1:z1"), xlEdgeTop, xlHairline

        Set dealWithUpClause6 = upClause6RangObj

End Function


Private Function dealWithUpClause7(upClause7RangObj As Range, sourceDataAsDicUpIssuingStatus As Object) As Variant
        'put all the information clause 7 on new UP

        Dim valueQtySummaryAsUPDict As Object
        Set valueQtySummaryAsUPDict = Application.Run("createUp.valueQtySummaryAsUP", sourceDataAsDicUpIssuingStatus)

        'delete all the lc information rows
        upClause7RangObj.Rows("2:" & upClause7RangObj.Rows.Count - 1).EntireRow.Delete

        'insert rows as lc count
        Dim i As Long
        For i = 1 To sourceDataAsDicUpIssuingStatus.Count * 2
            upClause7RangObj.Rows("2").EntireRow.Insert
        Next i

        upClause7RangObj.Range("r" & upClause7RangObj.Rows.Count).value = valueQtySummaryAsUPDict("TotalExportQtySumConvertedInYds")
        upClause7RangObj.Range("t" & upClause7RangObj.Rows.Count).value = valueQtySummaryAsUPDict("TotalExportValueSumConvertedInUSD")

        Application.Run "utility_formating_fun.SetBorderInsideHairlineAroundThin", upClause7RangObj.Range("b1:aa" & upClause7RangObj.Rows.Count)

        Dim j As Long
        Dim tempRange As Range

        Dim lcKey As Variant


        Dim fontName, fontSize, rowsHeight As Variant

        fontName = "Arial Narrow"
        fontSize = 12
        rowsHeight = 42

        Application.DisplayAlerts = False

            'commented, cause no need to change header
        ' lcKey = sourceDataAsDicUpIssuingStatus.keys()(0) 'take first lc key
        ' Set tempRange = upClause7RangObj(1, 1).Resize(1, upClause7RangObj.Columns.Count)  'set header row
        ' Application.Run "createUp.putHeaderFieldAsFirstLcInfoUpClause7", tempRange, sourceDataAsDicUpIssuingStatus, lcKey


        For j = 1 To sourceDataAsDicUpIssuingStatus.Count

            lcKey = sourceDataAsDicUpIssuingStatus.keys()(j - 1)

            Set tempRange = upClause7RangObj(j * 2, 1).Resize(2, upClause7RangObj.Columns.Count)

            Application.Run "utility_formating_fun.rangeFormat", tempRange, fontName, fontSize, False, True, xlCenter, xlCenter, "General"
            tempRange.Range("a1:a2").EntireRow.RowHeight = rowsHeight

            'put sl. no.
            tempRange(1, 2) = j
            tempRange(1, 2).Resize(2).Merge

            Application.Run "createUp.putCommonFieldAsLcInfoUpClause7", tempRange, sourceDataAsDicUpIssuingStatus, lcKey

            Application.Run "createUp.putProductDescriptionFieldAsLcInfoUpClause7", tempRange, sourceDataAsDicUpIssuingStatus, lcKey

            Application.Run "createUp.putProductQtyFieldAsLcInfoUpClause7", tempRange, sourceDataAsDicUpIssuingStatus, lcKey

            Application.Run "createUp.putProductValueFieldAsLcInfoUpClause7", tempRange, sourceDataAsDicUpIssuingStatus, lcKey

            Application.Run "createUp.putIpExpMLcFieldAsLcInfoUpClause7", tempRange, sourceDataAsDicUpIssuingStatus, lcKey


        Next j

        Application.DisplayAlerts = True



        Set dealWithUpClause7 = upClause7RangObj

End Function

Private Function combinLcAndAmnd(lcDict As Object) As String

    Dim temp As String
    temp = lcDict("LCSCNo") & Chr(10) & DateValue(lcDict("LCIssueDate"))

    If Not IsEmpty(lcDict("BangladeshBankRef")) Then
    
        temp = temp & Chr(10) & "(DC-" & lcDict("BangladeshBankRef") & ")"
    
    End If
    
    If lcDict("LCAmndNo") <> "-" Then
        Dim amndNo As Variant
        amndNo = Application.Run("general_utility_functions.ExtractRightDigitFromEnd", lcDict("LCAmndNo"))   'take right digits from end
        If amndNo < 10 Then
            amndNo = "0" & amndNo
        End If
        temp = temp & Chr(10) & "Amnd-" & amndNo & " Dt." & lcDict("LCAmndDate")
    End If

    combinLcAndAmnd = temp
    
End Function

Private Function valueQtySummaryAsUP(sourceDataAsDicUpIssuingStatus As Object) As Object

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")

    resultDict("TotalExportValueSumConvertedInUSD") = 0
    resultDict("TotalExportQtySumConvertedInYds") = 0
    resultDict("totalB2BValueSum") = 0
    resultDict("totalB2BQtySum") = 0
    resultDict("sumOfGarmentsQty") = 0

    resultDict("exportValueSumOfUSD") = 0
    resultDict("exportValueSumOfEuro") = 0

    resultDict("exportQtySumOfYds") = 0
    resultDict("exportQtySumOfMtr") = 0

    resultDict("usedInUpB2BValueSum") = 0
    resultDict("notUsedInUpB2BValueSum") = 0

    resultDict("usedInUpB2BQtySum") = 0
    resultDict("notUsedInUpB2BQtySum") = 0

    Dim dicKey As Variant

    For Each dicKey In sourceDataAsDicUpIssuingStatus.keys

        resultDict("TotalExportValueSumConvertedInUSD") = resultDict("TotalExportValueSumConvertedInUSD") + Application.Run("createUp.valueInUsd", sourceDataAsDicUpIssuingStatus(dicKey))
        resultDict("TotalExportQtySumConvertedInYds") = resultDict("TotalExportQtySumConvertedInYds") + Application.Run("createUp.qtyInYds", sourceDataAsDicUpIssuingStatus(dicKey))
        resultDict("totalB2BValueSum") = resultDict("totalB2BValueSum") + sourceDataAsDicUpIssuingStatus(dicKey)("BTBAmount")
        resultDict("totalB2BQtySum") = resultDict("totalB2BQtySum") + sourceDataAsDicUpIssuingStatus(dicKey)("QuantityKgs")
        resultDict("sumOfGarmentsQty") = resultDict("sumOfGarmentsQty") + sourceDataAsDicUpIssuingStatus(dicKey)("GarmentsQty")

        If Left(sourceDataAsDicUpIssuingStatus(dicKey)("currencyNumberFormat"), 8) = vsCodeNotSupportedOrBengaliTxtDictionary("sourceDataAsDicUpIssuingStatusCurrencyNumberFormat") Then

            resultDict("exportValueSumOfEuro") = resultDict("exportValueSumOfEuro") + CDbl(sourceDataAsDicUpIssuingStatus(dicKey)("LCAmount"))

        Else

            resultDict("exportValueSumOfUSD") = resultDict("exportValueSumOfUSD") + CDbl(sourceDataAsDicUpIssuingStatus(dicKey)("LCAmount"))

        End If

        If Right(sourceDataAsDicUpIssuingStatus(dicKey)("qtyNumberFormat"), 5) = """Mtr""" Then

            resultDict("exportQtySumOfMtr") = resultDict("exportQtySumOfMtr") + sourceDataAsDicUpIssuingStatus(dicKey)("QuantityofFabricsYdsMtr")

        Else

            resultDict("exportQtySumOfYds") = resultDict("exportQtySumOfYds") + sourceDataAsDicUpIssuingStatus(dicKey)("QuantityofFabricsYdsMtr")

        End If

        If Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("b2bComment"), "B2B not use in UP", True, True, True) Then

            resultDict("notUsedInUpB2BValueSum") = resultDict("notUsedInUpB2BValueSum") + sourceDataAsDicUpIssuingStatus(dicKey)("BTBAmount")
            resultDict("notUsedInUpB2BQtySum") = resultDict("notUsedInUpB2BQtySum") + sourceDataAsDicUpIssuingStatus(dicKey)("QuantityKgs")

        Else

            resultDict("usedInUpB2BValueSum") = resultDict("usedInUpB2BValueSum") + sourceDataAsDicUpIssuingStatus(dicKey)("BTBAmount")
            resultDict("usedInUpB2BQtySum") = resultDict("usedInUpB2BQtySum") + sourceDataAsDicUpIssuingStatus(dicKey)("QuantityKgs")

        End If

    Next dicKey

    Set valueQtySummaryAsUP = resultDict
    
End Function

Private Function qtyInYds(lcDict As Object) As Variant

    Dim temp As Variant

    If Right(lcDict("qtyNumberFormat"), 5) = """Mtr""" Then

        temp = Round(lcDict("QuantityofFabricsYdsMtr") * 1.0936132983)

    Else

        temp = lcDict("QuantityofFabricsYdsMtr")

    End If

    qtyInYds = temp
    
End Function

Private Function valueInUsd(lcDict As Object) As Variant

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")
    Dim temp As Variant

    If Left(lcDict("currencyNumberFormat"), 8) = vsCodeNotSupportedOrBengaliTxtDictionary("sourceDataAsDicUpIssuingStatusCurrencyNumberFormat") Then

        temp = CDbl(Round(lcDict("LCAmount") * 1.05)) ' conversion rate would be dynamic

    Else

        temp = CDbl(lcDict("LCAmount"))

    End If

    valueInUsd = temp
    
End Function

Private Function putCommonFieldAsLcInfoUpClause7(lcRangObj As Range, sourceDataAsDicUpIssuingStatus As Object, lcKey As Variant)
    'fill-up common field as lc information

    'put LC no.
    lcRangObj(1, 3).NumberFormat = "@"
    lcRangObj(1, 3).value = Application.Run("createUp.combinLcAndAmnd", sourceDataAsDicUpIssuingStatus(lcKey))
    lcRangObj(1, 3).Resize(2, 7).Merge

    'put Bank
    lcRangObj(1, 10).value = sourceDataAsDicUpIssuingStatus(lcKey)("LCIssuingBank")
    lcRangObj(1, 10).Resize(2, 6).Merge

    'put shipment date
    lcRangObj(1, 16).value = DateValue(sourceDataAsDicUpIssuingStatus(lcKey)("ShipmentDate"))
    lcRangObj(1, 16).VerticalAlignment = xlBottom
    Application.Run "utility_formating_fun.removeBorder", lcRangObj(1, 16), xlEdgeBottom

    'put shipment date
    lcRangObj(2, 16).value = DateValue(sourceDataAsDicUpIssuingStatus(lcKey)("ExpiryDate"))
    lcRangObj(2, 16).VerticalAlignment = xlTop

End Function


Private Function putProductDescriptionFieldAsLcInfoUpClause7(lcRangObj As Range, sourceDataAsDicUpIssuingStatus As Object, lcKey As Variant)
    'this function fill-up product description field as lc information


    If IsEmpty(sourceDataAsDicUpIssuingStatus(lcKey)("GarmentsQty")) Then

        lcRangObj(1, 17).value = "Denim Fabric"
        lcRangObj(1, 17).Resize(2, 1).Merge

    Else

        lcRangObj(1, 17).value = "Denim Garments"
        lcRangObj(2, 17).value = "Denim Fabric"

    End If

End Function


Private Function putProductQtyFieldAsLcInfoUpClause7(lcRangObj As Range, sourceDataAsDicUpIssuingStatus As Object, lcKey As Variant)
    'this function fill-up product qty. field as lc information

    lcRangObj(1, 18).Resize(2, 2).Style = "Comma"

    If IsEmpty(sourceDataAsDicUpIssuingStatus(lcKey)("GarmentsQty")) Then

        If Right(sourceDataAsDicUpIssuingStatus(lcKey)("qtyNumberFormat"), 5) = """Mtr""" Then

            lcRangObj(1, 18).value = WorksheetFunction.Text(sourceDataAsDicUpIssuingStatus(lcKey)("QuantityofFabricsYdsMtr"), "#,##0.00") & " Mtr"
            lcRangObj(1, 18).Resize(1, 2).Merge
            lcRangObj(1, 18).HorizontalAlignment = xlRight
            lcRangObj(2, 18).value = Round(sourceDataAsDicUpIssuingStatus(lcKey)("QuantityofFabricsYdsMtr") * 1.0936132983)
            lcRangObj(2, 18).Resize(1, 2).Merge

        Else

            lcRangObj(1, 18).value = sourceDataAsDicUpIssuingStatus(lcKey)("QuantityofFabricsYdsMtr")
            lcRangObj(1, 18).Resize(2, 2).Merge

        End If


    Else

        lcRangObj(1, 18).value = sourceDataAsDicUpIssuingStatus(lcKey)("GarmentsQty")
        lcRangObj(1, 18).Resize(1, 2).Merge
        lcRangObj(2, 18).value = sourceDataAsDicUpIssuingStatus(lcKey)("QuantityofFabricsYdsMtr")
        lcRangObj(2, 18).Resize(1, 2).Merge

    End If

End Function

Private Function putProductValueFieldAsLcInfoUpClause7(lcRangObj As Range, sourceDataAsDicUpIssuingStatus As Object, lcKey As Variant)
    'this function fill-up product value field as lc information

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    lcRangObj(1, 20).Resize(2, 2).Style = "Comma"

    If Left(sourceDataAsDicUpIssuingStatus(lcKey)("currencyNumberFormat"), 8) = vsCodeNotSupportedOrBengaliTxtDictionary("sourceDataAsDicUpIssuingStatusCurrencyNumberFormat") Then

'        lcRangObj(1, 20).NumberFormat = "@"
        lcRangObj(1, 20).value = "Euro  " & WorksheetFunction.Text(sourceDataAsDicUpIssuingStatus(lcKey)("LCAmount"), "#,##0.00")
        lcRangObj(1, 20).Resize(1, 2).Merge
        lcRangObj(1, 20).HorizontalAlignment = xlRight
        lcRangObj(2, 20).value = CDbl(Round(sourceDataAsDicUpIssuingStatus(lcKey)("LCAmount") * 1.05)) ' conversion rate would be dynamic
        lcRangObj(2, 20).Resize(1, 2).Merge

    Else

        lcRangObj(1, 20).value = CDbl(sourceDataAsDicUpIssuingStatus(lcKey)("LCAmount"))
        lcRangObj(1, 20).Resize(2, 2).Merge

    End If

End Function


Private Function putIpExpMLcFieldAsLcInfoUpClause7(lcRangObj As Range, sourceDataAsDicUpIssuingStatus As Object, lcKey As Variant)
    'fill-up Ip or Exp or M.LC field as lc information

    lcRangObj(1, 22).value = Application.Run("createUp.combinUdIpExpMlc", sourceDataAsDicUpIssuingStatus(lcKey), 10, 10)
    lcRangObj(1, 22).Resize(2, 6).Merge

End Function

Private Function combinUdIpExpMlc(lcDict As Object, innerConcateCharacterCode As Long, afterDateConcateCharacterCode As Long) As String

    Dim udIpExp As Object
    Set udIpExp = Application.Run("general_utility_functions.sequentiallyRelateTwoArraysAsDictionary", "udOrIpOrExp", "date", Split(lcDict("UDNoIPNo"), Chr(10)), Split(lcDict("UDIPDate"), Chr(10)))

    Dim mLC As Object
    Set mLC = Application.Run("general_utility_functions.sequentiallyRelateTwoArraysAsDictionary", "mLcNo", "date", Split(lcDict("MasterLCNo"), Chr(10)), Split(lcDict("MasterLCIssueDt"), Chr(10)))

    Dim concateExp As String
    Dim concateIp As String
    Dim concateMLc As String
    Dim returnStr As String

    If IsEmpty(lcDict("GarmentsQty")) Then

        If Application.Run("general_utility_functions.isStrPatternExist", lcDict("UDNoIPNo"), "^IP", True, True, True) Then
            ' non Garments EPZ

            concateExp = Application.Run("createUp.udIpExpMLcWithDtFilterAndConcate", udIpExp, "^EXP", "udOrIpOrExp", "date", innerConcateCharacterCode, afterDateConcateCharacterCode)
            concateIp = Application.Run("createUp.udIpExpMLcWithDtFilterAndConcate", udIpExp, "^IP", "udOrIpOrExp", "date", innerConcateCharacterCode, afterDateConcateCharacterCode)
            returnStr = concateExp & Chr(afterDateConcateCharacterCode) & concateIp

        ElseIf Application.Run("general_utility_functions.isStrPatternExist", lcDict("UDNoIPNo"), "^EXP", True, True, True) Then
            ' non Garments direct

            concateExp = Application.Run("createUp.udIpExpMLcWithDtFilterAndConcate", udIpExp, "^EXP", "udOrIpOrExp", "date", innerConcateCharacterCode, afterDateConcateCharacterCode)
            returnStr = concateExp


        Else
            ' non Garments Deem

            concateMLc = Application.Run("createUp.udIpExpMLcWithDtFilterAndConcate", mLC, ".", "mLcNo", "date", innerConcateCharacterCode, afterDateConcateCharacterCode)
            returnStr = Trim(concateMLc)

        End If

    Else
        ' Garments
        returnStr = lcDict("LCSCNo") & Chr(innerConcateCharacterCode) & lcDict("LCIssueDate") ' just use LC or SC no. as MLC

    End If

    combinUdIpExpMlc = returnStr

End Function

Private Function combinUdIpExpAndDt(lcDict As Object, innerConcateCharacterCode As Long, afterDateConcateCharacterCode As Long) As String

    Dim udIpExp As Object
    Set udIpExp = Application.Run("general_utility_functions.sequentiallyRelateTwoArraysAsDictionary", "udOrIpOrExp", "date", Split(lcDict("UDNoIPNo"), Chr(10)), Split(lcDict("UDIPDate"), Chr(10)))

    Dim concateExp As String
    Dim concateIp As String
    Dim concateUd As String
    Dim returnStr As String


        If Application.Run("general_utility_functions.isStrPatternExist", lcDict("UDNoIPNo"), "^IP", True, True, True) Then
            ' EPZ
            concateExp = Application.Run("createUp.udIpExpMLcWithDtFilterAndConcate", udIpExp, "^EXP", "udOrIpOrExp", "date", innerConcateCharacterCode, afterDateConcateCharacterCode)
            concateIp = Application.Run("createUp.udIpExpMLcWithDtFilterAndConcate", udIpExp, "^IP", "udOrIpOrExp", "date", innerConcateCharacterCode, afterDateConcateCharacterCode)
            returnStr = concateExp & Chr(afterDateConcateCharacterCode) & concateIp

        ElseIf Application.Run("general_utility_functions.isStrPatternExist", lcDict("UDNoIPNo"), "^EXP", True, True, True) Then
            ' direct

            concateExp = Application.Run("createUp.udIpExpMLcWithDtFilterAndConcate", udIpExp, "^EXP", "udOrIpOrExp", "date", innerConcateCharacterCode, afterDateConcateCharacterCode)
            returnStr = concateExp


        Else
            ' Deem
            concateUd = Application.Run("createUp.udIpExpMLcWithDtFilterAndConcate", udIpExp, ".", "udOrIpOrExp", "date", innerConcateCharacterCode, afterDateConcateCharacterCode)
            returnStr = Trim(concateUd)

        End If

    combinUdIpExpAndDt = returnStr

End Function


Private Function udIpExpMLcWithDtFilterAndConcate(udIpExpMLcWithDtDic As Object, filterPattern As String, innerDicNameKey As String, innerDicDateKey As String, innerConcateCharacterCode As Long, afterDateConcateCharacterCode As Long)
     'this function take Ud, Ip, Exp or M.LC No. & Date dictionary, filter pattern, inner dictionary name key and inner dictionary date key, inner joinning character code then
     'return with concate sequentially Related Ud, Ip, Exp or M.LC No. & Date

    Dim dictKey As Variant
    Dim concateUdIpExpOrMLcWithDt As String
    Dim i As Long
    For i = 0 To udIpExpMLcWithDtDic.Count - 1

        dictKey = udIpExpMLcWithDtDic.keys()(i)

        If Application.Run("general_utility_functions.isStrPatternExist", dictKey, filterPattern, True, True, True) Then

            concateUdIpExpOrMLcWithDt = concateUdIpExpOrMLcWithDt & Trim(udIpExpMLcWithDtDic(dictKey)(innerDicNameKey)) & Chr(innerConcateCharacterCode) & Trim(udIpExpMLcWithDtDic(dictKey)(innerDicDateKey)) & Chr(afterDateConcateCharacterCode)

        End If

    Next i

    udIpExpMLcWithDtFilterAndConcate = Left(concateUdIpExpOrMLcWithDt, Len(concateUdIpExpOrMLcWithDt) - 1)

End Function


Private Function putHeaderFieldAsFirstLcInfoUpClause7(headerRangObj As Range, sourceDataAsDicUpIssuingStatus As Object, lcKey As Variant)
    'this function fill-up header row field as first lc information

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    headerRangObj(1, 22).Resize(1, 6).ClearContents
    headerRangObj(1, 22).Resize(1, 6).UnMerge

    If IsEmpty(sourceDataAsDicUpIssuingStatus(lcKey)("GarmentsQty")) Then

        If Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(lcKey)("UDNoIPNo"), "^IP", True, True, True) Then
            ' non Garments EPZ

            headerRangObj(1, 22).value = vsCodeNotSupportedOrBengaliTxtDictionary("expNoAndDtBengaliTxt") 'EXP
            headerRangObj(1, 22).Resize(1, 3).Merge

            headerRangObj(1, 25).value = "AvBwc bs I ZvwiL" 'IP
            headerRangObj(1, 25).Resize(1, 3).Merge

        ElseIf Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(lcKey)("UDNoIPNo"), "^EXP", True, True, True) Then
            ' non Garments direct

            headerRangObj(1, 22).value = vsCodeNotSupportedOrBengaliTxtDictionary("expNoAndDtBengaliTxt") 'EXP
            headerRangObj(1, 22).Resize(1, 6).Merge

        Else
            ' non Garments Deem

            headerRangObj(1, 22).value = vsCodeNotSupportedOrBengaliTxtDictionary("mlcNoAndDtBengaliTxt") 'MLc
            headerRangObj(1, 22).Resize(1, 6).Merge

        End If

    Else
        ' Garments

        headerRangObj(1, 22).value = vsCodeNotSupportedOrBengaliTxtDictionary("mlcNoAndDtBengaliTxt") 'MLc
        headerRangObj(1, 22).Resize(1, 6).Merge

    End If

End Function




