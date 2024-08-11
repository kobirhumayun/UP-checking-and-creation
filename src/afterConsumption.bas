Attribute VB_Name = "afterConsumption"
Option Explicit


Private Function upClause8MakeUniqueRowsFromProvidedWs(ws As Worksheet)

    Dim topRow As Variant
    topRow = ws.Cells.Find("8|  Avg`vwb Gjwmi weeiY t", LookAt:=xlPart).Row + 3

    Dim i As Long

    For i = topRow To 1000

        If IsEmpty(ws.Cells(i, 14).value) Then

            Exit For

        End If

        If ws.Cells(i, 14).value = ws.Cells(i, 14).Offset(-1, 0) Then

            ws.Cells(i, 14).EntireRow.Delete
            i = i - 1

        End If

    Next i

End Function


Private Function upYarnConsumptionInformationFromProvidedWs(ws As Worksheet) As Object
    'this function give yarn consumption information from consumption sheet

    Dim topRow, bottomRow As Variant
    Dim rgInfo As Variant

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    topRow = ws.Cells.Find("TOTAL", LookAt:=xlPart).Row - 4
    bottomRow = topRow + 16

    Dim workingRange As Range
    Set workingRange = ws.Range("C" & topRow & ":" & "K" & bottomRow)

    rgInfo = workingRange.value

    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(1, 1), rgInfo(1, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(2, 1), rgInfo(2, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(3, 1), rgInfo(3, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(5, 1), rgInfo(5, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(7, 1), rgInfo(7, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(8, 1), rgInfo(8, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(9, 1), rgInfo(9, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(10, 1), rgInfo(10, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(11, 1), rgInfo(11, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(12, 1), rgInfo(12, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(13, 1), rgInfo(13, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(14, 1), rgInfo(14, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(15, 1), rgInfo(15, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(16, 1), rgInfo(16, 8))
    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, rgInfo(17, 1), rgInfo(17, 8))

    Set upYarnConsumptionInformationFromProvidedWs = dict


End Function



Private Function upClause8InformationForCreateUpFromProvidedWs(ws As Worksheet, impPerformanceDataDic As Object) As Object
    'this function give source data as dictionary from UP clause8

    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find("8|  Avg`vwb Gjwmi weeiY t", LookAt:=xlPart).Row + 3
    bottomRow = ws.Range("V" & topRow).End(xlDown).Row - 1

    Dim workingRange As Range
    Set workingRange = ws.Range("A" & topRow & ":" & "AA" & bottomRow)

    Dim temp As Variant
    temp = workingRange.value

    Dim upClause8Dic As Object
    Set upClause8Dic = CreateObject("Scripting.Dictionary")

    Dim tempMushakOrBillOfEntryDic As Object

    Dim tempMuOrBillKey As String

    Dim propertiesArr, propertiesValArr As Variant

    ReDim propertiesArr(1 To 16)
    ReDim propertiesValArr(1 To 16)

    propertiesArr(1) = "lcNoAndDt"
    propertiesArr(2) = "mushakOrBillOfEntryNoAndDt"
    propertiesArr(3) = "nameOfGoods"
    propertiesArr(4) = "hsCode"
    propertiesArr(5) = "qtyOfGoods"
    propertiesArr(6) = "valueOfGoods"
    propertiesArr(7) = "previousUsedQtyOfGoods"
    propertiesArr(8) = "previousUsedValueOfGoods"
    propertiesArr(9) = "currentStockQtyOfGoods"
    propertiesArr(10) = "currentStockValueOfGoods"
    propertiesArr(11) = "inThisUpUsedQtyOfGoods"
    propertiesArr(12) = "inThisUpUsedValueOfGoods"
    propertiesArr(13) = "totalUsedQtyOfGoods"
    propertiesArr(14) = "totalUsedValueOfGoods"
    propertiesArr(15) = "remainingQtyOfGoods"
    propertiesArr(16) = "remainingValueOfGoods"

    Dim i As Long

    For i = 1 To UBound(temp) ' create dictionary as mushak or bill of entry

        propertiesValArr(1) = temp(i, 3)
        propertiesValArr(2) = temp(i, 7)
        propertiesValArr(3) = temp(i, 14)
        propertiesValArr(4) = temp(i, 15)
        propertiesValArr(5) = temp(i, 16)
        propertiesValArr(6) = temp(i, 17)
        propertiesValArr(7) = temp(i, 18)
        propertiesValArr(8) = temp(i, 19)
        propertiesValArr(9) = temp(i, 20)
        propertiesValArr(10) = temp(i, 21)
        propertiesValArr(11) = temp(i, 22)
        propertiesValArr(12) = temp(i, 23)
        propertiesValArr(13) = temp(i, 24)
        propertiesValArr(14) = temp(i, 25)
        propertiesValArr(15) = temp(i, 26)
        propertiesValArr(16) = temp(i, 27)

        Set tempMushakOrBillOfEntryDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

        If Not workingRange(i, 22).Comment Is Nothing Then   'check if the cell has a comment
            tempMushakOrBillOfEntryDic("inThisUpUsedQtyOfGoodsComment") = workingRange(i, 22).Comment.Text
        Else
            tempMushakOrBillOfEntryDic("inThisUpUsedQtyOfGoodsComment") = "No Comment"
        End If

        tempMuOrBillKey = Application.Run("general_utility_functions.dictKeyGeneratorWithLcMushakOrBillOfEntryQtyAndValue", temp(i, 3), temp(i, 7), temp(i, 16), temp(i, 17))

        upClause8Dic.Add tempMuOrBillKey, tempMushakOrBillOfEntryDic

    Next i

    Dim dicKey As Variant
    Dim upClause8DicGroupByGoods As Object
    Set upClause8DicGroupByGoods = CreateObject("Scripting.Dictionary")

    Dim removedAllInvalidChrFromRawMaterialsDes As String

    Dim yarnGroupNameDic As Object
    Set yarnGroupNameDic = impPerformanceDataDic("yarnGroupNameDic")

    For Each dicKey In upClause8Dic.keys

        removedAllInvalidChrFromRawMaterialsDes = Application.Run("general_utility_functions.RemoveInvalidChars", upClause8Dic(dicKey)("nameOfGoods"))   'remove all invalid characters

        If removedAllInvalidChrFromRawMaterialsDes = "Yarn" Then

            If Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("inThisUpUsedQtyOfGoodsComment"), "garments", True, True, True) Then

                removedAllInvalidChrFromRawMaterialsDes = "garments"

            ElseIf Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("inThisUpUsedQtyOfGoodsComment"), "cotton", True, True, True) Then

                removedAllInvalidChrFromRawMaterialsDes = "cotton"

            ElseIf Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("inThisUpUsedQtyOfGoodsComment"), "polyester", True, True, True) Then

                removedAllInvalidChrFromRawMaterialsDes = "polyester"

            ElseIf Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("inThisUpUsedQtyOfGoodsComment"), "spandex", True, True, True) Then

                removedAllInvalidChrFromRawMaterialsDes = "spandex"

            Else

                If yarnGroupNameDic.Exists(dicKey) Then

                    removedAllInvalidChrFromRawMaterialsDes = Application.Run("general_utility_functions.RemoveInvalidChars", yarnGroupNameDic(dicKey))    'remove all invalid characters

                Else

                    MsgBox dicKey & " not found in import performance"
                    Exit Function

                End If

            End If

            If Not upClause8DicGroupByGoods.Exists(removedAllInvalidChrFromRawMaterialsDes) Then ' create group by goods dictionary

                upClause8DicGroupByGoods.Add removedAllInvalidChrFromRawMaterialsDes, CreateObject("Scripting.Dictionary")

            End If

            upClause8DicGroupByGoods(removedAllInvalidChrFromRawMaterialsDes).Add dicKey, upClause8Dic(dicKey)

        Else

            If Not upClause8DicGroupByGoods.Exists(removedAllInvalidChrFromRawMaterialsDes) Then ' create group by goods dictionary

                upClause8DicGroupByGoods.Add removedAllInvalidChrFromRawMaterialsDes, CreateObject("Scripting.Dictionary")

            End If

            upClause8DicGroupByGoods(removedAllInvalidChrFromRawMaterialsDes).Add dicKey, upClause8Dic(dicKey)

        End If

    Next dicKey

    Set upClause8InformationForCreateUpFromProvidedWs = upClause8DicGroupByGoods

End Function


Private Function sourceDataAsDicImportYarnUseDetailsForUd(fileName As String, worksheetTabName As String) As Variant  ' Source file name & worksheetTabName
    'this function give source data as dictionary from Import Yarn Use Details For UD File

    Application.Run "utilityFunction.openFile", fileName ' provide filename

    ActiveWorkbook.Worksheets(worksheetTabName).Activate
    ActiveSheet.AutoFilterMode = False

    Dim workingRange As Range
    Set workingRange = Range("C4:" & "T" & Range("C4").End(xlDown).Row)

    Dim temp As Variant
    temp = workingRange.value

    Application.Run "utilityFunction.closeFile", fileName ' provide filename

    Dim importYarnUseDetailsGroupbyUdDic As Object
    Set importYarnUseDetailsGroupbyUdDic = CreateObject("Scripting.Dictionary")

    Dim tempMushakOrBillOfEntryDic As Object

    Dim tempMuOrBillKey As String

    Dim udNo As Variant

    Dim dicKey As Variant

    Dim propertiesArr, propertiesValArr As Variant

    ReDim propertiesArr(1 To 18)
    ReDim propertiesValArr(1 To 18)

    propertiesArr(1) = "lcNoAndDt"
    propertiesArr(2) = "mushakOrBillOfEntryNoAndDt"
    propertiesArr(3) = "nameOfGoods"
    propertiesArr(4) = "hsCode"
    propertiesArr(5) = "qtyOfGoods"
    propertiesArr(6) = "valueOfGoods"
    propertiesArr(7) = "previousUsedQtyOfGoods"
    propertiesArr(8) = "previousUsedValueOfGoods"
    propertiesArr(9) = "currentStockQtyOfGoods"
    propertiesArr(10) = "currentStockValueOfGoods"
    propertiesArr(11) = "inThisUpUsedQtyOfGoods" ' use UP instead of UD, for exception handle
    propertiesArr(12) = "inThisUpUsedValueOfGoods" ' use UP instead of UD, for exception handle
    propertiesArr(13) = "totalUsedQtyOfGoods"
    propertiesArr(14) = "totalUsedValueOfGoods"
    propertiesArr(15) = "remainingQtyOfGoods"
    propertiesArr(16) = "remainingValueOfGoods"
    propertiesArr(17) = "udNo"
    propertiesArr(18) = "udDate"

    Dim i As Long

    For i = 1 To UBound(temp) ' create dictionary as UD

        propertiesValArr(1) = temp(i, 2)
        propertiesValArr(2) = temp(i, 1)
        propertiesValArr(3) = temp(i, 4)
        propertiesValArr(4) = temp(i, 3)
        propertiesValArr(5) = temp(i, 5)
        propertiesValArr(6) = temp(i, 6)
        propertiesValArr(7) = 0 'no need because UP sequence not same as UD
        propertiesValArr(8) = 0 'no need because UP sequence not same as UD
        propertiesValArr(9) = 0 'no need because UP sequence not same as UD
        propertiesValArr(10) = 0 'no need because UP sequence not same as UD
        propertiesValArr(11) = temp(i, 11)
        propertiesValArr(12) = temp(i, 12)
        propertiesValArr(13) = 0 'no need because UP sequence not same as UD
        propertiesValArr(14) = 0 'no need because UP sequence not same as UD
        propertiesValArr(15) = 0 'no need because UP sequence not same as UD
        propertiesValArr(16) = 0 'no need because UP sequence not same as UD
        propertiesValArr(17) = temp(i, 17)
        propertiesValArr(18) = temp(i, 18)

        Set tempMushakOrBillOfEntryDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

        tempMuOrBillKey = Application.Run("general_utility_functions.dictKeyGeneratorWithLcMushakOrBillOfEntryQtyAndValue", temp(i, 2), temp(i, 1), temp(i, 5), temp(i, 6))


        Set udNo = Application.Run("general_utility_functions.regExReturnedObj", tempMushakOrBillOfEntryDic("udNo"), "\d+", True, True, True)

        If udNo.Count = 2 Then

            udNo = Val(udNo(0)) & "_" & Val(udNo(1))

        ElseIf udNo.Count = 3 Then

            udNo = Val(udNo(0)) & "_" & Val(udNo(1)) & "_" & Val(udNo(2))

        End If

        If Not importYarnUseDetailsGroupbyUdDic.Exists(udNo) Then ' create group by ud dictionary

            importYarnUseDetailsGroupbyUdDic.Add udNo, CreateObject("Scripting.Dictionary")

        End If

        importYarnUseDetailsGroupbyUdDic(udNo).Add tempMuOrBillKey, tempMushakOrBillOfEntryDic


    Next i


    Set sourceDataAsDicImportYarnUseDetailsForUd = importYarnUseDetailsGroupbyUdDic


End Function


Private Function createNewUpClause8Information(upClause8InfoDic As Object, impPerformanceDataDic As Object, sourceDataAsDicUpIssuingStatus As Object, importYarnUseDetailsForUd As Object, finalRawMaterialsQtyDicAsGroup As Object) As Object
    'this function create new UP clause8 data dictionary


    Dim newUpClause8InfoDic As Object
    Set newUpClause8InfoDic = CreateObject("Scripting.Dictionary")

    Dim newUpClause8TempYarnInfoDic As Object ' only for yarn sorting use this dict
    Set newUpClause8TempYarnInfoDic = CreateObject("Scripting.Dictionary")

    Dim upClause8KeysDic As Object
    Set upClause8KeysDic = CreateObject("Scripting.Dictionary")

    Dim usedB2bInfoFromUpIssuingStatus As Object
    Set usedB2bInfoFromUpIssuingStatus = CreateObject("Scripting.Dictionary")

    Dim allMushakInfoAgainstB2b As Object
    Set allMushakInfoAgainstB2b = CreateObject("Scripting.Dictionary")

    Dim allMushakInfoAgainstB2bAsUpClause8Format As Object
    Set allMushakInfoAgainstB2bAsUpClause8Format = CreateObject("Scripting.Dictionary")

    Dim allBillOfEntryInfoUsedInUd As Object
    Set allBillOfEntryInfoUsedInUd = CreateObject("Scripting.Dictionary")


    Dim dicKey As Variant
    Dim innerDicKey As Variant
    Dim innerDicKey1 As Variant

    Dim goodsGroupNameFromUpWs As Variant

    Dim tempDic As Object

    Dim isGarments As Variant
    
    If sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(0))("GarmentsQty") > 0 Then
    
        isGarments = True
    
    
    Else
    
        isGarments = False
    
    End If


    For Each dicKey In upClause8InfoDic.keys

        For Each innerDicKey In upClause8InfoDic(dicKey).keys ' take all mushak or bill of entry keys to check if already exist any unused qty.

            upClause8KeysDic.Add innerDicKey, innerDicKey

        Next innerDicKey

    Next dicKey


    For Each dicKey In sourceDataAsDicUpIssuingStatus.keys

        If IsDate(sourceDataAsDicUpIssuingStatus(dicKey)("BTBLCIssueDate")) And Not Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("b2bComment"), "B2B not use in UP", True, True, True) Then

            Set tempDic = CreateObject("Scripting.Dictionary")

            tempDic("BTBLCNo") = sourceDataAsDicUpIssuingStatus(dicKey)("BTBLCNo")
            tempDic("BTBLCIssueDate") = sourceDataAsDicUpIssuingStatus(dicKey)("BTBLCIssueDate")
            tempDic("BTBAmount") = sourceDataAsDicUpIssuingStatus(dicKey)("BTBAmount")
            tempDic("QuantityKgs") = sourceDataAsDicUpIssuingStatus(dicKey)("QuantityKgs")
            tempDic("b2bComment") = sourceDataAsDicUpIssuingStatus(dicKey)("b2bComment")

            usedB2bInfoFromUpIssuingStatus.Add sourceDataAsDicUpIssuingStatus(dicKey)("BTBLCNo"), tempDic

        End If

    Next dicKey


    For Each dicKey In usedB2bInfoFromUpIssuingStatus.keys

        If impPerformanceDataDic("CottonYarnLocalOrImpClassifiedDbDic")("localCtnAsLc").Exists(dicKey) Then

            Set allMushakInfoAgainstB2b = Application.Run("dictionary_utility_functions.mergeDict", allMushakInfoAgainstB2b, impPerformanceDataDic("CottonYarnLocalOrImpClassifiedDbDic")("localCtnAsLc")(dicKey))

        Else

            ' properties take from "CombinedAllSheetsMushakOrBillOfEntryDbDict" Function, if properties mismatch than arises conflict
            Set tempDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", _
            Array("BillOfEntryOrMushak", "LC", "HSCode", "Description", "Qty", "Value", "UsedQty", "UsedValue", "BalanceQty", "BalanceValue"), _
            Array(usedB2bInfoFromUpIssuingStatus(dicKey)("BTBLCNo"), usedB2bInfoFromUpIssuingStatus(dicKey)("BTBLCNo"), "5203.00.00", "COTTON YARN", _
            usedB2bInfoFromUpIssuingStatus(dicKey)("QuantityKgs"), usedB2bInfoFromUpIssuingStatus(dicKey)("BTBAmount"), _
            0, 0, _
            usedB2bInfoFromUpIssuingStatus(dicKey)("QuantityKgs"), usedB2bInfoFromUpIssuingStatus(dicKey)("BTBAmount")))

            allMushakInfoAgainstB2b.Add dicKey, tempDic ' if mushak not found then add lc info from up issuing status

        End If

    Next dicKey


    For Each dicKey In allMushakInfoAgainstB2b.keys

        allMushakInfoAgainstB2bAsUpClause8Format.Add dicKey, Application.Run("afterConsumption.convertImpPerformanceMushakOrBillOfEntryToUpClause8", allMushakInfoAgainstB2b(dicKey))

    Next dicKey


    If isGarments Then

        Dim formatedUdNo As String

        For Each dicKey In sourceDataAsDicUpIssuingStatus.keys ' specific bill of entry use in UD

            formatedUdNo = Application.Run("general_utility_functions.extractAndFormatUdNo", sourceDataAsDicUpIssuingStatus(dicKey)("UDNoIPNo"))

            If importYarnUseDetailsForUd.Exists(formatedUdNo) Then

                Set tempDic = importYarnUseDetailsForUd(formatedUdNo)

            Else

                Set tempDic = CreateObject("Scripting.Dictionary")

            End If

            For Each innerDicKey In tempDic.keys

                If Not allBillOfEntryInfoUsedInUd.Exists(innerDicKey) Then

                    allBillOfEntryInfoUsedInUd.Add innerDicKey, tempDic(innerDicKey)

                Else

                    allBillOfEntryInfoUsedInUd(innerDicKey)("inThisUpUsedQtyOfGoods") = allBillOfEntryInfoUsedInUd(innerDicKey)("inThisUpUsedQtyOfGoods") + tempDic(innerDicKey)("inThisUpUsedQtyOfGoods")

                End If

            Next innerDicKey

        Next dicKey

    End If


    newUpClause8InfoDic.Add "Yarn", CreateObject("Scripting.Dictionary") 'just add dict for take top position "newUpClause8InfoDic" dict

    For Each dicKey In upClause8InfoDic.keys


        If dicKey = "garments" Or dicKey = "cotton" Or dicKey = "polyester" Or dicKey = "spandex" Then

            If isGarments Then

                If dicKey = "garments" Then

                    If allMushakInfoAgainstB2bAsUpClause8Format.Count > 0 Then

                        Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", allMushakInfoAgainstB2bAsUpClause8Format, upClause8KeysDic, CreateObject("Scripting.Dictionary"), finalRawMaterialsQtyDicAsGroup("cotton")) ' cotton comsumption use only B2B, no add any Bill of entry

                        newUpClause8TempYarnInfoDic.Add "garmentsCotton", tempDic

                    End If

                    Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", upClause8InfoDic(dicKey), upClause8KeysDic, CreateObject("Scripting.Dictionary"), 0) ' use 0, no add any Bill of entry, just for others calculation

                    For Each innerDicKey In allBillOfEntryInfoUsedInUd.keys

                        If tempDic.Exists(innerDicKey) Then ' for specific bill of entry use in UD

                            tempDic(innerDicKey)("inThisUpUsedQtyOfGoods") = allBillOfEntryInfoUsedInUd(innerDicKey)("inThisUpUsedQtyOfGoods")

                        Else

                            If allBillOfEntryInfoUsedInUd(innerDicKey)("previousUsedQtyOfGoods") = 0 Then

                                allBillOfEntryInfoUsedInUd(innerDicKey)("inThisUpUsedQtyOfGoodsComment") = "ONLY USED FOR GARMENTS UD" ' add comment for new garments Bill of Entry

                                tempDic.Add innerDicKey, allBillOfEntryInfoUsedInUd(innerDicKey)

                            Else

                                MsgBox "Not inserted new Bill Of Entry " & innerDicKey & " in UP due to previous Qty. not 0 in UD"

                            End If

                        End If

                    Next innerDicKey

                    newUpClause8TempYarnInfoDic.Add "garments", tempDic

                Else

                    Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", upClause8InfoDic(dicKey), upClause8KeysDic, CreateObject("Scripting.Dictionary"), 0) ' for non garments use "0" cause garments up

                    newUpClause8TempYarnInfoDic.Add dicKey, tempDic


                End If


            Else ' non garments


                If dicKey = "garments" Then

                    Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", upClause8InfoDic(dicKey), upClause8KeysDic, CreateObject("Scripting.Dictionary"), 0) ' for garments use "0" cause non garments up

                    newUpClause8TempYarnInfoDic.Add dicKey, tempDic

                Else

                    If dicKey = "cotton" Then

                        If allMushakInfoAgainstB2bAsUpClause8Format.Count > 0 Then

                            Set tempDic = Application.Run("dictionary_utility_functions.mergeDict", allMushakInfoAgainstB2bAsUpClause8Format, upClause8InfoDic(dicKey))

                        Else

                            Set tempDic = upClause8InfoDic(dicKey)

                        End If


                        If impPerformanceDataDic("CottonYarnLocalOrImpClassifiedDbDic")("importCtnAsBillOfEntry").Count > 0 Then

                            Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", tempDic, upClause8KeysDic, impPerformanceDataDic("CottonYarnLocalOrImpClassifiedDbDic")("importCtnAsBillOfEntry"), finalRawMaterialsQtyDicAsGroup(dicKey))

                        Else

                            Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", tempDic, upClause8KeysDic, CreateObject("Scripting.Dictionary"), finalRawMaterialsQtyDicAsGroup(dicKey))

                        End If

                        newUpClause8TempYarnInfoDic.Add dicKey, tempDic

                    Else

                        ' "impPerformanceDataDic("yarnClassifiedDbDic")" this dict contain "cotton", "polyester" & "spandex"
                        ' but never take "cotton", cause enter if block when dicKey = "cotton"
                        ' note "cotton" devided into import and local category which handle in if block


                        If impPerformanceDataDic("yarnClassifiedDbDic").Exists(dicKey) Then

                            Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", upClause8InfoDic(dicKey), upClause8KeysDic, impPerformanceDataDic("yarnClassifiedDbDic")(dicKey), finalRawMaterialsQtyDicAsGroup(dicKey))

                        Else

                            Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", upClause8InfoDic(dicKey), upClause8KeysDic, CreateObject("Scripting.Dictionary"), finalRawMaterialsQtyDicAsGroup(dicKey))

                        End If

                        newUpClause8TempYarnInfoDic.Add dicKey, tempDic

                    End If

                End If

            End If

        Else


            If impPerformanceDataDic("nonYarnClassifiedDbDic").Exists(dicKey) Then

                    'name of goods take from first mushak or bill of entry from previous UP, which mean keep original group name from UP worksheet
                goodsGroupNameFromUpWs = upClause8InfoDic(dicKey)(upClause8InfoDic(dicKey).keys()(0))("nameOfGoods")

                Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", upClause8InfoDic(dicKey), upClause8KeysDic, impPerformanceDataDic("nonYarnClassifiedDbDic")(dicKey), finalRawMaterialsQtyDicAsGroup(dicKey))

            Else

                    'name of goods take from first mushak or bill of entry from previous UP, which mean keep original group name from UP worksheet
                goodsGroupNameFromUpWs = upClause8InfoDic(dicKey)(upClause8InfoDic(dicKey).keys()(0))("nameOfGoods")

                Set tempDic = Application.Run("afterConsumption.createNewUpClause8DicGroupByGoods", upClause8InfoDic(dicKey), upClause8KeysDic, CreateObject("Scripting.Dictionary"), finalRawMaterialsQtyDicAsGroup(dicKey))

            End If
            
                'change name of goods from UP worksheet's group name
            For Each innerDicKey1 In tempDic.keys

                tempDic(innerDicKey1)("nameOfGoods") = goodsGroupNameFromUpWs

            Next innerDicKey1

            newUpClause8InfoDic.Add dicKey, tempDic


        End If


    Next dicKey

        'change name of goods to "Yarn"
    For Each dicKey In newUpClause8TempYarnInfoDic.keys

        For Each innerDicKey In newUpClause8TempYarnInfoDic(dicKey).keys

            newUpClause8TempYarnInfoDic(dicKey)(innerDicKey)("nameOfGoods") = "Yarn"

        Next innerDicKey

    Next dicKey

    ' yarn dict key sorting only

    If newUpClause8TempYarnInfoDic.Exists("garmentsCotton") Then

        Set newUpClause8InfoDic("Yarn") = Application.Run("dictionary_utility_functions.mergeDict", newUpClause8InfoDic("Yarn"), newUpClause8TempYarnInfoDic("garmentsCotton"))

    End If

    If newUpClause8TempYarnInfoDic.Exists("cotton") Then

        Set newUpClause8InfoDic("Yarn") = Application.Run("dictionary_utility_functions.mergeDict", newUpClause8InfoDic("Yarn"), newUpClause8TempYarnInfoDic("cotton"))

    End If

    If newUpClause8TempYarnInfoDic.Exists("polyester") Then

        Set newUpClause8InfoDic("Yarn") = Application.Run("dictionary_utility_functions.mergeDict", newUpClause8InfoDic("Yarn"), newUpClause8TempYarnInfoDic("polyester"))

    End If

    If newUpClause8TempYarnInfoDic.Exists("spandex") Then

        Set newUpClause8InfoDic("Yarn") = Application.Run("dictionary_utility_functions.mergeDict", newUpClause8InfoDic("Yarn"), newUpClause8TempYarnInfoDic("spandex"))

    End If

    If newUpClause8TempYarnInfoDic.Exists("garments") Then

        Set newUpClause8InfoDic("Yarn") = Application.Run("dictionary_utility_functions.mergeDict", newUpClause8InfoDic("Yarn"), newUpClause8TempYarnInfoDic("garments"))

    End If

    Set createNewUpClause8Information = newUpClause8InfoDic

End Function


Private Function createNewUpClause8DicGroupByGoods(upClause8DicGroupByGoods As Object, upClause8KeysDic As Object, groupByGoodsFromimpPerformance As Object, totalUseQty As Variant) As Object
    'this function create new UP clause8 data dictionary

    Dim newUpClause8InfoDic As Object
    Set newUpClause8InfoDic = CreateObject("Scripting.Dictionary")

    Dim tempMushakOrBillOfEntryDic As Object

    Dim remainingQtySum As Variant

    Dim dicKey As Variant
    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = upClause8DicGroupByGoods(upClause8DicGroupByGoods.keys()(0)).keys 'take properties for dynamic

    ReDim propertiesValArr(0 To 16) ' declared arr 0-16(not 1-17) for matching with properties

    remainingQtySum = 0

    For Each dicKey In upClause8DicGroupByGoods.keys

        remainingQtySum = remainingQtySum + upClause8DicGroupByGoods(dicKey)("remainingQtyOfGoods")

    Next dicKey


    For Each dicKey In groupByGoodsFromimpPerformance.keys

        If remainingQtySum < totalUseQty Then

            If Not upClause8KeysDic.Exists(dicKey) Then 'exclude(handle) if already inserted clause 8 but not use

                propertiesValArr(0) = groupByGoodsFromimpPerformance(dicKey)("LC")
                propertiesValArr(1) = groupByGoodsFromimpPerformance(dicKey)("BillOfEntryOrMushak")
                propertiesValArr(2) = groupByGoodsFromimpPerformance(dicKey)("Description")
                propertiesValArr(3) = groupByGoodsFromimpPerformance(dicKey)("HSCode")
                propertiesValArr(4) = groupByGoodsFromimpPerformance(dicKey)("Qty")
                propertiesValArr(5) = groupByGoodsFromimpPerformance(dicKey)("Value")
                propertiesValArr(6) = 0
                propertiesValArr(7) = 0
                propertiesValArr(8) = 0
                propertiesValArr(9) = 0
                propertiesValArr(10) = 0
                propertiesValArr(11) = 0
                propertiesValArr(12) = 0
                propertiesValArr(13) = 0
                propertiesValArr(14) = groupByGoodsFromimpPerformance(dicKey)("Qty") ' remaining Qty. of goods is same of Qty.
                propertiesValArr(15) = groupByGoodsFromimpPerformance(dicKey)("Value") ' remaining value of goods is same of value
                propertiesValArr(16) = "No Comment"

                Set tempMushakOrBillOfEntryDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)


                upClause8DicGroupByGoods.Add dicKey, tempMushakOrBillOfEntryDic
                remainingQtySum = remainingQtySum + tempMushakOrBillOfEntryDic("remainingQtyOfGoods")

            End If


        Else

            Exit For

        End If


    Next dicKey


    For Each dicKey In upClause8DicGroupByGoods.keys ' add to new dict only remained qty. greater than 0

        upClause8DicGroupByGoods(dicKey)("previousUsedQtyOfGoods") = upClause8DicGroupByGoods(dicKey)("previousUsedQtyOfGoods") + upClause8DicGroupByGoods(dicKey)("inThisUpUsedQtyOfGoods")
        upClause8DicGroupByGoods(dicKey)("previousUsedValueOfGoods") = upClause8DicGroupByGoods(dicKey)("previousUsedValueOfGoods") + upClause8DicGroupByGoods(dicKey)("inThisUpUsedValueOfGoods")

        upClause8DicGroupByGoods(dicKey)("inThisUpUsedQtyOfGoods") = 0

        If upClause8DicGroupByGoods(dicKey)("remainingQtyOfGoods") > 0.1 Then

            newUpClause8InfoDic.Add dicKey, upClause8DicGroupByGoods(dicKey)

        End If

    Next dicKey


    For Each dicKey In newUpClause8InfoDic.keys ' use calculation in this UP


        If newUpClause8InfoDic(dicKey)("remainingQtyOfGoods") <= totalUseQty Then

            newUpClause8InfoDic(dicKey)("inThisUpUsedQtyOfGoods") = newUpClause8InfoDic(dicKey)("remainingQtyOfGoods")
            totalUseQty = totalUseQty - newUpClause8InfoDic(dicKey)("inThisUpUsedQtyOfGoods")

        Else

            newUpClause8InfoDic(dicKey)("inThisUpUsedQtyOfGoods") = totalUseQty
            Exit For

        End If

    Next dicKey


    If newUpClause8InfoDic.Count = 0 Then ' if not exist any remained qty. greater than 0
        dicKey = upClause8DicGroupByGoods.keys()(0)
        newUpClause8InfoDic.Add dicKey, upClause8DicGroupByGoods(dicKey)
    End If


    Set createNewUpClause8DicGroupByGoods = newUpClause8InfoDic

End Function



Private Function upClause8InformationPutToProvidedWs(ws As Worksheet, newUpClause8Dic As Object)
    'this function put new UP clause8 information to provided worksheet

    Dim upClause8DicGroupByGoods As Object

    Dim topRow As Variant
    topRow = ws.Cells.Find("8|  Avg`vwb Gjwmi weeiY t", LookAt:=xlPart).Row + 3

    Dim i, j, loopCounter, mushakOrBillOfEntrySl As Long

    Dim dicKey As Variant

    mushakOrBillOfEntrySl = 0

    For i = topRow To 1000

        If IsEmpty(ws.Cells(i, 14).value) Then

            Exit For

        End If

        Set upClause8DicGroupByGoods = newUpClause8Dic(Application.Run("general_utility_functions.RemoveInvalidChars", ws.Cells(i, 14).value)) ' handle yarn here finally

        'insert rows as mushak or bill of entry count, note already one row exist
        If upClause8DicGroupByGoods.Count > 1 Then

            For j = 1 To upClause8DicGroupByGoods.Count - 1
                ws.Cells(i, 14).Rows("2").EntireRow.Insert
            Next j

        End If

        loopCounter = 0

        For Each dicKey In upClause8DicGroupByGoods.keys

            mushakOrBillOfEntrySl = mushakOrBillOfEntrySl + 1

            ws.Cells(i + loopCounter, 2).value = mushakOrBillOfEntrySl
            ws.Cells(i + loopCounter, 3).value = upClause8DicGroupByGoods(dicKey)("lcNoAndDt")
            ws.Cells(i + loopCounter, 7).value = upClause8DicGroupByGoods(dicKey)("mushakOrBillOfEntryNoAndDt")
            ws.Cells(i + loopCounter, 14).value = upClause8DicGroupByGoods(dicKey)("nameOfGoods")
            ws.Cells(i + loopCounter, 15).value = upClause8DicGroupByGoods(dicKey)("hsCode")
            ws.Cells(i + loopCounter, 16).value = upClause8DicGroupByGoods(dicKey)("qtyOfGoods")
            ws.Cells(i + loopCounter, 17).value = upClause8DicGroupByGoods(dicKey)("valueOfGoods")
            ws.Cells(i + loopCounter, 18).value = upClause8DicGroupByGoods(dicKey)("previousUsedQtyOfGoods")
            ws.Cells(i + loopCounter, 19).value = upClause8DicGroupByGoods(dicKey)("previousUsedValueOfGoods")
            ws.Cells(i + loopCounter, 20).value = upClause8DicGroupByGoods(dicKey)("currentStockQtyOfGoods")
            ws.Cells(i + loopCounter, 21).value = upClause8DicGroupByGoods(dicKey)("currentStockValueOfGoods")
            ws.Cells(i + loopCounter, 22).value = upClause8DicGroupByGoods(dicKey)("inThisUpUsedQtyOfGoods")
            ws.Cells(i + loopCounter, 23).value = upClause8DicGroupByGoods(dicKey)("inThisUpUsedValueOfGoods")
            ws.Cells(i + loopCounter, 24).value = upClause8DicGroupByGoods(dicKey)("totalUsedQtyOfGoods")
            ws.Cells(i + loopCounter, 25).value = upClause8DicGroupByGoods(dicKey)("totalUsedValueOfGoods")
            ws.Cells(i + loopCounter, 26).value = upClause8DicGroupByGoods(dicKey)("remainingQtyOfGoods")
            ws.Cells(i + loopCounter, 27).value = upClause8DicGroupByGoods(dicKey)("remainingValueOfGoods")


            ws.Cells(i + loopCounter, 22).ClearComments

            If upClause8DicGroupByGoods(dicKey)("inThisUpUsedQtyOfGoodsComment") <> "No Comment" Then

                ws.Cells(i + loopCounter, 22).AddComment upClause8DicGroupByGoods(dicKey)("inThisUpUsedQtyOfGoodsComment")

            End If


            ws.Cells(i + loopCounter, 20).FormulaR1C1 = "=RC[-4]-RC[-2]"
            ws.Cells(i + loopCounter, 21).FormulaR1C1 = "=RC[-4]-RC[-2]"

                ' handle error, if divide zero then error show in cell
            If upClause8DicGroupByGoods(dicKey)("qtyOfGoods") > 0 Then

                ws.Cells(i + loopCounter, 23).FormulaR1C1 = "=RC[-6]/RC[-7]*RC[-1]"

            Else

                ws.Cells(i + loopCounter, 23).value = 0

            End If
            
            ws.Cells(i + loopCounter, 24).FormulaR1C1 = "=SUM(RC[-6],RC[-2])"
            ws.Cells(i + loopCounter, 25).FormulaR1C1 = "=SUM(RC[-6],RC[-2])"
            ws.Cells(i + loopCounter, 26).FormulaR1C1 = "=RC[-6]-RC[-4]"
            ws.Cells(i + loopCounter, 27).FormulaR1C1 = "=RC[-6]-RC[-4]"
            ws.Cells(i + loopCounter, 28).FormulaR1C1 = "=RC[-11]/RC[-12]"
            ws.Cells(i + loopCounter, 29).FormulaR1C1 = "=RC[-10]/RC[-11]"
            ws.Cells(i + loopCounter, 30).FormulaR1C1 = "=RC[-9]/RC[-10]"
            ws.Cells(i + loopCounter, 31).FormulaR1C1 = "=RC[-8]/RC[-9]"
            ws.Cells(i + loopCounter, 32).FormulaR1C1 = "=RC[-7]/RC[-8]"
            ws.Cells(i + loopCounter, 33).FormulaR1C1 = "=RC[-6]/RC[-7]"


            ws.Cells(i + loopCounter, 3).Resize(1, 4).Merge
            ws.Cells(i + loopCounter, 7).Resize(1, 7).Merge


            loopCounter = loopCounter + 1


        Next dicKey


        i = i + upClause8DicGroupByGoods.Count - 1

    Next i

    ws.Cells(i, 22).FormulaR1C1 = "=SUM(R[-" & i - topRow & "]C:R[-1]C)"
    ws.Cells(i, 23).FormulaR1C1 = "=SUM(R[-" & i - topRow & "]C:R[-1]C)"


    Application.Run "utility_formating_fun.SetBorderInsideHairlineAroundThin", ws.Range(ws.Cells(topRow - 2, 2), ws.Cells(i, 27))

    ws.Range(ws.Cells(topRow, 16), ws.Cells(i, 27)).Style = "Comma"


End Function



Private Function convertImpPerformanceMushakOrBillOfEntryToUpClause8(mushakOrBillOfEntryFromImpPerformance As Object) As Object
    'this function convert import performance single mushak or bill of entry to up clause 8 mushak or bill of entry

    Dim propertiesArr, propertiesValArr As Variant
    Dim tempMushakOrBillOfEntryDic As Object

    ReDim propertiesArr(1 To 17)
    ReDim propertiesValArr(1 To 17)


    ' properties take from "upClause8InformationForCreateUpFromProvidedWs" Function, if properties mismatch than arises conflict

    propertiesArr(1) = "lcNoAndDt"
    propertiesArr(2) = "mushakOrBillOfEntryNoAndDt"
    propertiesArr(3) = "nameOfGoods"
    propertiesArr(4) = "hsCode"
    propertiesArr(5) = "qtyOfGoods"
    propertiesArr(6) = "valueOfGoods"
    propertiesArr(7) = "previousUsedQtyOfGoods"
    propertiesArr(8) = "previousUsedValueOfGoods"
    propertiesArr(9) = "currentStockQtyOfGoods"
    propertiesArr(10) = "currentStockValueOfGoods"
    propertiesArr(11) = "inThisUpUsedQtyOfGoods"
    propertiesArr(12) = "inThisUpUsedValueOfGoods"
    propertiesArr(13) = "totalUsedQtyOfGoods"
    propertiesArr(14) = "totalUsedValueOfGoods"
    propertiesArr(15) = "remainingQtyOfGoods"
    propertiesArr(16) = "remainingValueOfGoods"
    propertiesArr(17) = "inThisUpUsedQtyOfGoodsComment"


    propertiesValArr(1) = mushakOrBillOfEntryFromImpPerformance("LC")
    propertiesValArr(2) = mushakOrBillOfEntryFromImpPerformance("BillOfEntryOrMushak")
    propertiesValArr(3) = mushakOrBillOfEntryFromImpPerformance("Description")
    propertiesValArr(4) = mushakOrBillOfEntryFromImpPerformance("HSCode")
    propertiesValArr(5) = mushakOrBillOfEntryFromImpPerformance("Qty")
    propertiesValArr(6) = mushakOrBillOfEntryFromImpPerformance("Value")
    propertiesValArr(7) = 0
    propertiesValArr(8) = 0
    propertiesValArr(9) = 0
    propertiesValArr(10) = 0
    propertiesValArr(11) = 0
    propertiesValArr(12) = 0
    propertiesValArr(13) = 0
    propertiesValArr(14) = 0
    propertiesValArr(15) = mushakOrBillOfEntryFromImpPerformance("Qty") ' remaining Qty. of goods is same of Qty.
    propertiesValArr(16) = mushakOrBillOfEntryFromImpPerformance("Value") ' remaining value of goods is same of value
    propertiesValArr(17) = "No Comment"

    Set tempMushakOrBillOfEntryDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set convertImpPerformanceMushakOrBillOfEntryToUpClause8 = tempMushakOrBillOfEntryDic

End Function

Private Function dealWithUpClause9(ws As Worksheet, newUpClause8InfoClassifiedPartDic As Object, sourceDataImportPerformanceTotalSummary As Variant)

    Dim upClause9StockinformationRangeObject As Variant
    Dim upClause9Val As Variant
    Dim temp As Variant

    Set upClause9StockinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause9StockinformationRangeObjectFromProvidedWs", ws)
    Set upClause9StockinformationRangeObject = upClause9StockinformationRangeObject(1, 1).Resize(6, 29)

    upClause9Val = upClause9StockinformationRangeObject.Value
    ReDim temp(1 To UBound(upClause9Val, 1), 1 To 1)

    Dim i As Long

        'previous used Qty. update
    For i = 1 To UBound(upClause9Val, 1)
        temp(i, 1) = upClause9Val(i, 28)
    Next i

    upClause9StockinformationRangeObject.Columns(20) = temp

        'new import Qty. update
    For i = 1 To UBound(upClause9Val, 1)
        temp(i, 1) = sourceDataImportPerformanceTotalSummary(i + 1, 5)
    Next i

    upClause9StockinformationRangeObject.Columns(16) = temp

        'used in this UP Qty. update
        temp(1, 1) = newUpClause8InfoClassifiedPartDic("yarnImportQty")
        temp(2, 1) = newUpClause8InfoClassifiedPartDic("yarnLocalQty")
        temp(3, 1) = newUpClause8InfoClassifiedPartDic("dyesQty")
        temp(4, 1) = newUpClause8InfoClassifiedPartDic("chemicalsImportQty")
        temp(5, 1) = newUpClause8InfoClassifiedPartDic("chemicalsLocalQty")
        temp(6, 1) = newUpClause8InfoClassifiedPartDic("stretchWrappingFilmQty")

    upClause9StockinformationRangeObject.Columns(24) = temp

End Function

Private Function dealWithUpClause9WithPreviousUpData(ws As Worksheet, newUpClause8InfoClassifiedPartDic As Object, sourceDataImportPerformanceTotalSummary As Variant, previousUpClause9Val As Variant)

    Dim upClause9StockinformationRangeObject As Variant
    Dim upClause9Val As Variant
    Dim temp As Variant

    Set upClause9StockinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause9StockinformationRangeObjectFromProvidedWs", ws)
    Set upClause9StockinformationRangeObject = upClause9StockinformationRangeObject(1, 1).Resize(6, 29)

    upClause9Val = upClause9StockinformationRangeObject.Value
    ReDim temp(1 To UBound(upClause9Val, 1), 1 To 1)

    Dim i As Long

        'previous used Qty. update
    For i = 1 To UBound(previousUpClause9Val, 1)
        temp(i, 1) = previousUpClause9Val(i, 28)
    Next i

    upClause9StockinformationRangeObject.Columns(20) = temp

        'new import Qty. update
    For i = 1 To UBound(upClause9Val, 1)
        temp(i, 1) = sourceDataImportPerformanceTotalSummary(i + 1, 5)
    Next i

    upClause9StockinformationRangeObject.Columns(16) = temp

        'used in this UP Qty. update
        temp(1, 1) = newUpClause8InfoClassifiedPartDic("yarnImportQty")
        temp(2, 1) = newUpClause8InfoClassifiedPartDic("yarnLocalQty")
        temp(3, 1) = newUpClause8InfoClassifiedPartDic("dyesQty")
        temp(4, 1) = newUpClause8InfoClassifiedPartDic("chemicalsImportQty")
        temp(5, 1) = newUpClause8InfoClassifiedPartDic("chemicalsLocalQty")
        temp(6, 1) = newUpClause8InfoClassifiedPartDic("stretchWrappingFilmQty")

    upClause9StockinformationRangeObject.Columns(24) = temp

End Function

Private Function upClause9InfoFromProvidedFile(upfileName As String, openFile As Boolean, closeFile As Boolean) As Variant
    
    If openFile Then
        Application.Run "utilityFunction.openFile", upfileName ' provide filename
    End If
    
    ActiveWorkbook.Worksheets(2).Activate
    ActiveSheet.AutoFilterMode = False

    Dim upClause9StockinformationRangeObject As Object
    Set upClause9StockinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause9StockinformationRangeObjectFromProvidedWs", ActiveSheet)
    Set upClause9StockinformationRangeObject = upClause9StockinformationRangeObject(1, 1).Resize(6, 29)
    
    Dim upClause9Val As Variant
    upClause9Val = upClause9StockinformationRangeObject.Value
    
    If closeFile Then
       Application.Run "utilityFunction.closeFile", upfileName ' provide filename
    End If
    
    upClause9InfoFromProvidedFile = upClause9Val

End Function

Private Function sumNewUpClause8ClassifiedPart(newUpClause8Dic As Object) As Object
    
    Dim dicKey As Variant
    Dim innerDicKey As Variant
    Dim calculatedValue As Variant

    Dim YarnDyesChemicalsClassifiedPart As Object
    Set YarnDyesChemicalsClassifiedPart = CreateObject("Scripting.Dictionary")

        'assign zero to handle if any actual Qty. zero, then put empty string
    YarnDyesChemicalsClassifiedPart("yarnImportQty") = 0
    YarnDyesChemicalsClassifiedPart("yarnImportValue") = 0
    YarnDyesChemicalsClassifiedPart("yarnLocalQty") = 0
    YarnDyesChemicalsClassifiedPart("yarnLocalValue") = 0
    YarnDyesChemicalsClassifiedPart("dyesQty") = 0
    YarnDyesChemicalsClassifiedPart("dyesValue") = 0
    YarnDyesChemicalsClassifiedPart("stretchWrappingFilmQty") = 0
    YarnDyesChemicalsClassifiedPart("stretchWrappingFilmValue") = 0
    YarnDyesChemicalsClassifiedPart("chemicalsImportQty") = 0
    YarnDyesChemicalsClassifiedPart("chemicalsImportValue") = 0
    YarnDyesChemicalsClassifiedPart("chemicalsLocalQty") = 0
    YarnDyesChemicalsClassifiedPart("chemicalsLocalValue") = 0

    For Each dicKey In newUpClause8Dic.keys

        For Each innerDicKey In newUpClause8Dic(dicKey).keys

                'calculate value, cause used value not calculated, used value to be calculate in excel sheet
            If newUpClause8Dic(dicKey)(innerDicKey)("valueOfGoods") > 0 Then

                calculatedValue = newUpClause8Dic(dicKey)(innerDicKey)("valueOfGoods") / newUpClause8Dic(dicKey)(innerDicKey)("qtyOfGoods") * newUpClause8Dic(dicKey)(innerDicKey)("inThisUpUsedQtyOfGoods")

            Else

                calculatedValue = 0

            End If

            If Application.Run("general_utility_functions.isStrPatternExist", newUpClause8Dic(dicKey)(innerDicKey)("nameOfGoods"), "yarn", True, True, True) Then

                If Application.Run("general_utility_functions.isStrPatternExist", newUpClause8Dic(dicKey)(innerDicKey)("mushakOrBillOfEntryNoAndDt"), "^c-", True, True, True) Then

                    YarnDyesChemicalsClassifiedPart("yarnImportQty") = YarnDyesChemicalsClassifiedPart("yarnImportQty") + newUpClause8Dic(dicKey)(innerDicKey)("inThisUpUsedQtyOfGoods")
                    YarnDyesChemicalsClassifiedPart("yarnImportValue") = YarnDyesChemicalsClassifiedPart("yarnImportValue") + calculatedValue

                ElseIf Application.Run("general_utility_functions.isStrPatternExist", newUpClause8Dic(dicKey)(innerDicKey)("mushakOrBillOfEntryNoAndDt"), "^m", True, True, True) Then
                    
                    YarnDyesChemicalsClassifiedPart("yarnLocalQty") = YarnDyesChemicalsClassifiedPart("yarnLocalQty") + newUpClause8Dic(dicKey)(innerDicKey)("inThisUpUsedQtyOfGoods")
                    YarnDyesChemicalsClassifiedPart("yarnLocalValue") = YarnDyesChemicalsClassifiedPart("yarnLocalValue") + calculatedValue

                End If

            ElseIf Application.Run("general_utility_functions.isStrPatternExist", newUpClause8Dic(dicKey)(innerDicKey)("nameOfGoods"), "dyes", True, True, True)  Then

                YarnDyesChemicalsClassifiedPart("dyesQty") = YarnDyesChemicalsClassifiedPart("dyesQty") + newUpClause8Dic(dicKey)(innerDicKey)("inThisUpUsedQtyOfGoods")
                YarnDyesChemicalsClassifiedPart("dyesValue") = YarnDyesChemicalsClassifiedPart("dyesValue") + calculatedValue

            ElseIf Application.Run("general_utility_functions.isStrPatternExist", newUpClause8Dic(dicKey)(innerDicKey)("nameOfGoods"), "Stretch Wrapping Film", True, True, True)  Then
            
                YarnDyesChemicalsClassifiedPart("stretchWrappingFilmQty") = YarnDyesChemicalsClassifiedPart("stretchWrappingFilmQty") + newUpClause8Dic(dicKey)(innerDicKey)("inThisUpUsedQtyOfGoods")
                YarnDyesChemicalsClassifiedPart("stretchWrappingFilmValue") = YarnDyesChemicalsClassifiedPart("stretchWrappingFilmValue") + calculatedValue

            Else

                If Application.Run("general_utility_functions.isStrPatternExist", newUpClause8Dic(dicKey)(innerDicKey)("mushakOrBillOfEntryNoAndDt"), "^c-", True, True, True) Then

                    YarnDyesChemicalsClassifiedPart("chemicalsImportQty") = YarnDyesChemicalsClassifiedPart("chemicalsImportQty") + newUpClause8Dic(dicKey)(innerDicKey)("inThisUpUsedQtyOfGoods")
                    YarnDyesChemicalsClassifiedPart("chemicalsImportValue") = YarnDyesChemicalsClassifiedPart("chemicalsImportValue") + calculatedValue

                ElseIf Application.Run("general_utility_functions.isStrPatternExist", newUpClause8Dic(dicKey)(innerDicKey)("mushakOrBillOfEntryNoAndDt"), "^m", True, True, True) Then

                    YarnDyesChemicalsClassifiedPart("chemicalsLocalQty") = YarnDyesChemicalsClassifiedPart("chemicalsLocalQty") + newUpClause8Dic(dicKey)(innerDicKey)("inThisUpUsedQtyOfGoods")
                    YarnDyesChemicalsClassifiedPart("chemicalsLocalValue") = YarnDyesChemicalsClassifiedPart("chemicalsLocalValue") + calculatedValue

                End If
                
            End If

        Next innerDicKey

    Next dicKey


    Set sumNewUpClause8ClassifiedPart = YarnDyesChemicalsClassifiedPart

End Function

Private Function sumUpClause8ClassifiedPart(upClause8Dic As Object) As Object
    
    Dim dicKey As Variant

    Dim YarnDyesChemicalsClassifiedPart As Object
    Set YarnDyesChemicalsClassifiedPart = CreateObject("Scripting.Dictionary")

        'assign zero to handle if any actual Qty. zero, then put empty string
    YarnDyesChemicalsClassifiedPart("yarnImportQty") = 0
    YarnDyesChemicalsClassifiedPart("yarnImportValue") = 0
    YarnDyesChemicalsClassifiedPart("yarnLocalQty") = 0
    YarnDyesChemicalsClassifiedPart("yarnLocalValue") = 0
    YarnDyesChemicalsClassifiedPart("dyesQty") = 0
    YarnDyesChemicalsClassifiedPart("dyesValue") = 0
    YarnDyesChemicalsClassifiedPart("stretchWrappingFilmQty") = 0
    YarnDyesChemicalsClassifiedPart("stretchWrappingFilmValue") = 0
    YarnDyesChemicalsClassifiedPart("chemicalsImportQty") = 0
    YarnDyesChemicalsClassifiedPart("chemicalsImportValue") = 0
    YarnDyesChemicalsClassifiedPart("chemicalsLocalQty") = 0
    YarnDyesChemicalsClassifiedPart("chemicalsLocalValue") = 0

    For Each dicKey In upClause8Dic.keys

            If Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("nameOfGoods"), "yarn", True, True, True) Then

                If Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("mushakOrBillOfEntryNoAndDt"), "^c-", True, True, True) Then

                    YarnDyesChemicalsClassifiedPart("yarnImportQty") = YarnDyesChemicalsClassifiedPart("yarnImportQty") + upClause8Dic(dicKey)("inThisUpUsedQtyOfGoods")
                    YarnDyesChemicalsClassifiedPart("yarnImportValue") = YarnDyesChemicalsClassifiedPart("yarnImportValue") + upClause8Dic(dicKey)("inThisUpUsedValueOfGoods")

                ElseIf Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("mushakOrBillOfEntryNoAndDt"), "^m", True, True, True) Then
                    
                    YarnDyesChemicalsClassifiedPart("yarnLocalQty") = YarnDyesChemicalsClassifiedPart("yarnLocalQty") + upClause8Dic(dicKey)("inThisUpUsedQtyOfGoods")
                    YarnDyesChemicalsClassifiedPart("yarnLocalValue") = YarnDyesChemicalsClassifiedPart("yarnLocalValue") + upClause8Dic(dicKey)("inThisUpUsedValueOfGoods")

                End If

            ElseIf Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("nameOfGoods"), "dyes", True, True, True)  Then

                YarnDyesChemicalsClassifiedPart("dyesQty") = YarnDyesChemicalsClassifiedPart("dyesQty") + upClause8Dic(dicKey)("inThisUpUsedQtyOfGoods")
                YarnDyesChemicalsClassifiedPart("dyesValue") = YarnDyesChemicalsClassifiedPart("dyesValue") + upClause8Dic(dicKey)("inThisUpUsedValueOfGoods")

            ElseIf Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("nameOfGoods"), "Stretch Wrapping Film", True, True, True)  Then
            
                YarnDyesChemicalsClassifiedPart("stretchWrappingFilmQty") = YarnDyesChemicalsClassifiedPart("stretchWrappingFilmQty") + upClause8Dic(dicKey)("inThisUpUsedQtyOfGoods")
                YarnDyesChemicalsClassifiedPart("stretchWrappingFilmValue") = YarnDyesChemicalsClassifiedPart("stretchWrappingFilmValue") + upClause8Dic(dicKey)("inThisUpUsedValueOfGoods")

            Else

                If Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("mushakOrBillOfEntryNoAndDt"), "^c-", True, True, True) Then

                    YarnDyesChemicalsClassifiedPart("chemicalsImportQty") = YarnDyesChemicalsClassifiedPart("chemicalsImportQty") + upClause8Dic(dicKey)("inThisUpUsedQtyOfGoods")
                    YarnDyesChemicalsClassifiedPart("chemicalsImportValue") = YarnDyesChemicalsClassifiedPart("chemicalsImportValue") + upClause8Dic(dicKey)("inThisUpUsedValueOfGoods")

                ElseIf Application.Run("general_utility_functions.isStrPatternExist", upClause8Dic(dicKey)("mushakOrBillOfEntryNoAndDt"), "^m", True, True, True) Then

                    YarnDyesChemicalsClassifiedPart("chemicalsLocalQty") = YarnDyesChemicalsClassifiedPart("chemicalsLocalQty") + upClause8Dic(dicKey)("inThisUpUsedQtyOfGoods")
                    YarnDyesChemicalsClassifiedPart("chemicalsLocalValue") = YarnDyesChemicalsClassifiedPart("chemicalsLocalValue") + upClause8Dic(dicKey)("inThisUpUsedValueOfGoods")

                End If
                
            End If

    Next dicKey

    Set sumUpClause8ClassifiedPart = YarnDyesChemicalsClassifiedPart

End Function

Private Function dealWithUpClause11(ws As Worksheet, sourceDataAsDicUpIssuingStatus As Object)

    Dim upClause11UdExpIpinformationRangeObject As Range
    Set upClause11UdExpIpinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause11UdExpIpinformationRangeObjectFromProvidedWs", ws)

    If upClause11UdExpIpinformationRangeObject.Rows.Count > 2 Then

        upClause11UdExpIpinformationRangeObject.Rows("2:" & upClause11UdExpIpinformationRangeObject.Rows.Count - 1).EntireRow.Delete

    End If
    
    Dim i, j, k, l, m As Long

    'insert rows as lc count, note already two rows exist one row for UD, IP, EXP, buyer etc info and one row form total sum
    'rest row insert between these rows
    If sourceDataAsDicUpIssuingStatus.Count > 1 Then

        For i = 1 To sourceDataAsDicUpIssuingStatus.Count - 1
            upClause11UdExpIpinformationRangeObject.Rows("2").EntireRow.Insert
        Next i

    End If

    Dim regExReturnedObjectUdIpExp As Object
    Dim regExReturnedObjectUdIpExpDt As Object
    Dim tempUdIpExpAndDtJoinStr As String

    Dim tempWidthStr As Object
    Dim tempWeightStr As Object
    Dim dicKey As Variant

    Dim temp As Variant

    For j = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

            'put Sl. No.
        upClause11UdExpIpinformationRangeObject.Range("b" & j + 1).value = j + 1
        ' upClause11UdExpIpinformationRangeObject.Range("b" & j + 1 & ":c" & j + 1).Merge

            'put buyer name
        upClause11UdExpIpinformationRangeObject.Range("c" & j + 1).value = sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("NameofBuyers")
        upClause11UdExpIpinformationRangeObject.Range("c" & j + 1 & ":p" & j + 1).Merge


            'extract UD IP EXP & Date
        Set regExReturnedObjectUdIpExp = Application.Run("general_utility_functions.regExReturnedObj", sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("UDNoIPNo"), ".+", True, True, True)
        Set regExReturnedObjectUdIpExpDt = Application.Run("general_utility_functions.regExReturnedObj", sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("UDIPDate"), ".+", True, True, True)


        For k = 0 To regExReturnedObjectUdIpExp.Count - 1

                'join UD IP EXP & Date
            tempUdIpExpAndDtJoinStr = tempUdIpExpAndDtJoinStr & regExReturnedObjectUdIpExp(k) & " " & regExReturnedObjectUdIpExpDt(k) & Chr(10)

        Next k

            'put UD IP EXP & Date
        upClause11UdExpIpinformationRangeObject.Range("q" & j + 1).value = Left(tempUdIpExpAndDtJoinStr, Len(tempUdIpExpAndDtJoinStr) - 1)
        upClause11UdExpIpinformationRangeObject.Range("q" & j + 1 & ":s" & j + 1).Merge

        tempUdIpExpAndDtJoinStr = "" 'reset

        upClause11UdExpIpinformationRangeObject.Range("t" & j + 1).value = "Denim Garments"
        upClause11UdExpIpinformationRangeObject.Range("t" & j + 1 & ":u" & j + 1).Merge

        upClause11UdExpIpinformationRangeObject.Range("v" & j + 1).value = "Denim Fabric"

        Set tempWidthStr = CreateObject("Scripting.Dictionary")
        Set tempWeightStr = CreateObject("Scripting.Dictionary")

        For Each dicKey In sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange").keys

                'create unique width & weight
            tempWidthStr(sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange")(dicKey)("width").value) = Null
            tempWeightStr(sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange")(dicKey)("weight").value) = Null

        Next dicKey

            'sort
        temp = Application.Run("Sorting_Algorithms.BubbleSort", tempWidthStr.keys)

            'formate
        For l = LBound(temp) To UBound(temp)
            temp(l) = Format(temp(l), "0.00")
        Next l
        
            'put width
        upClause11UdExpIpinformationRangeObject.Range("w" & j + 1).value = Join(temp, ",")
        upClause11UdExpIpinformationRangeObject.Range("w" & j + 1 & ":x" & j + 1).Merge

            'sort
        temp = Application.Run("Sorting_Algorithms.BubbleSort", tempWeightStr.keys)
        
            'format
        For m = LBound(temp) To UBound(temp)
            temp(m) = Format(temp(m), "0.00")
        Next m
        
            'put weight
        upClause11UdExpIpinformationRangeObject.Range("y" & j + 1).value = Join(temp, ",")

            'put qty.
        If Right(sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("qtyNumberFormat"), 5) = """Mtr""" Then

            upClause11UdExpIpinformationRangeObject.Range("z" & j + 1).value = Round(sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("QuantityofFabricsYdsMtr") * 1.0936132983)
            upClause11UdExpIpinformationRangeObject.Range("z" & j + 1 & ":aa" & j + 1).Merge

        Else

            upClause11UdExpIpinformationRangeObject.Range("z" & j + 1).value = sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("QuantityofFabricsYdsMtr")
            upClause11UdExpIpinformationRangeObject.Range("z" & j + 1 & ":aa" & j + 1).Merge


        End If

    Next j

    upClause11UdExpIpinformationRangeObject.Range("z" & upClause11UdExpIpinformationRangeObject.Rows.Count).FormulaR1C1 = "=SUM(R[-" & upClause11UdExpIpinformationRangeObject.Rows.Count - 1 & "]C:R[-1]C)"

    Set upClause11UdExpIpinformationRangeObject = upClause11UdExpIpinformationRangeObject.Offset(-2, 1) _
    .Resize(upClause11UdExpIpinformationRangeObject.Rows.Count + 2, upClause11UdExpIpinformationRangeObject.Columns.Count - 1)

    Application.Run "utility_formating_fun.SetBorderInsideHairlineAroundThin", upClause11UdExpIpinformationRangeObject
    Application.Run "utility_formating_fun.removeBorder", upClause11UdExpIpinformationRangeObject.Range _
    ("a" & upClause11UdExpIpinformationRangeObject.Rows.Count & ":x" & upClause11UdExpIpinformationRangeObject.Rows.Count), xlInsideVertical

End Function

Private Function addConRangeToSourceDataAsDicUpIssuingStatus(ws As Worksheet, sourceDataAsDicUpIssuingStatus As Object)

    Dim i, j As Long
    Dim loopCounter As Long

    loopCounter = 4 'initially run from first buyer row

    For i = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

        j = loopCounter

        If sourceDataAsDicUpIssuingStatus.Count = 1 Then

            If ws.Cells(j, 1) = sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("NameofBuyers") Then
                ' Debug.Print "match"
                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i)).Add "consumptionRange", CreateObject("Scripting.Dictionary")
            Else
                ' Debug.Print "mismatch " & sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("NameofBuyers")
                MsgBox "mismatch " & sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("NameofBuyers") & " consumption sheet"
            End If

        Else

            If ws.Cells(j, 1) = i + 1 & ") " & sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("NameofBuyers") Then
                ' Debug.Print "match"
                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i)).Add "consumptionRange", CreateObject("Scripting.Dictionary")
            Else
                ' Debug.Print "mismatch " & i + 1 & ") " & sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("NameofBuyers")
                MsgBox "mismatch " & i + 1 & ") " & sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("NameofBuyers") & " consumption sheet"
            End If

        End If


        j = j + 1 ' add one for check next row, as no need to check buyer row

        Do Until ws.Cells(j, 3) = "Cotton"

        
            If ws.Cells(j, 1) = "Weight :" Then

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Add _
                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count + 1, CreateObject("Scripting.Dictionary")
                
                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "weight", ws.Cells(j, 4)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "width", ws.Cells(j, 12)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "qty", ws.Cells(j, 20)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "yarnConPerYds", ws.Cells(j + 2, 15)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "totalYarnCon", ws.Cells(j + 8, 3)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "cottonYarn", ws.Cells(j + 10, 6)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "polyesterYarn", ws.Cells(j + 11, 6)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "spandexYarn", ws.Cells(j + 12, 6)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "black", ws.Cells(j + 2, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "mercerizationBlack", ws.Cells(j + 3, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "indigo", ws.Cells(j + 4, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "mercerizationIndigo", ws.Cells(j + 5, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "toppingBottoming", ws.Cells(j + 6, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "mercerizationToppingBottoming", ws.Cells(j + 7, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "overDying", ws.Cells(j + 8, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "mercerizationOverDying", ws.Cells(j + 9, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "coating", ws.Cells(j + 10, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "pfd", ws.Cells(j + 11, 19)

                sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange") _
                (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("consumptionRange").Count).Add "ecru", ws.Cells(j + 12, 19)

                ' Debug.Print i + 1 & ") " & ws.Cells(j, 4)

            End If
            
            j = j + 1

            If j > 2000 Then ' asume highest consumption
                Exit Do
            End If

           If Not IsEmpty(ws.Cells(j, 1)) And ws.Cells(j, 1) <> "Cotton" And ws.Cells(j, 1) <> "Polyester" And ws.Cells(j, 1) <> "Spandex" And ws.Cells(j, 1) <> "Weight :" Then

                Exit Do ' asume in this row exist buyer

            End If
            
        Loop

        loopCounter = j

    Next i

    Set addConRangeToSourceDataAsDicUpIssuingStatus = sourceDataAsDicUpIssuingStatus

End Function

Private Function dealWithUpClause12a(ws As Worksheet, sourceDataAsDicUpIssuingStatus As Object)

    Dim upClause12AYarnConsumptionInformationRangeObject As Range
    Set upClause12AYarnConsumptionInformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause12AYarnConsumptioninformationRangeObjectFromProvidedWs", ws)

    If upClause12AYarnConsumptionInformationRangeObject.Rows.Count > 2 Then

        upClause12AYarnConsumptionInformationRangeObject.Rows("2:" & upClause12AYarnConsumptionInformationRangeObject.Rows.Count - 1).EntireRow.Delete

    End If
    
    Dim dicKey As Variant
    Dim consumptionCount As Long

    For Each dicKey In sourceDataAsDicUpIssuingStatus.keys

        consumptionCount = consumptionCount + sourceDataAsDicUpIssuingStatus(dicKey)("consumptionRange").Count

    Next dicKey

    Dim i, j, k, l, m As Long

    'insert rows as consumption count, note already two rows exist one row for consumption info and one row for total sum
    'rest row insert between these rows
    If consumptionCount > 1 Then

        For i = 1 To consumptionCount - 1
            upClause12AYarnConsumptionInformationRangeObject.Rows("2").EntireRow.Insert
        Next i

    End If

    Dim loopCounter As Long

    loopCounter = 1
    
    For j = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

            'put Sl. No.
        upClause12AYarnConsumptionInformationRangeObject.Range("b" & loopCounter).value = j + 1
        upClause12AYarnConsumptionInformationRangeObject.Range("b" & loopCounter & ":b" & loopCounter + sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange").Count - 1).Merge

            'put buyer name
        upClause12AYarnConsumptionInformationRangeObject.Range("c" & loopCounter).value = sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("NameofBuyers")
        upClause12AYarnConsumptionInformationRangeObject.Range("c" & loopCounter & ":o" & loopCounter + sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange").Count - 1).Merge

        For k = 0 To sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange").Count - 1

                'put goods description (fabrics)
            upClause12AYarnConsumptionInformationRangeObject.Range("p" & loopCounter + k).value = "Denim Fabrics"
            upClause12AYarnConsumptionInformationRangeObject.Range("p" & loopCounter + k & ":q" & loopCounter + k).Merge

                'put goods description (garments)
            upClause12AYarnConsumptionInformationRangeObject.Range("r" & loopCounter + k).value = 0

                'put yds
            upClause12AYarnConsumptionInformationRangeObject.Range("s" & loopCounter + k).value = sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange") _
            (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange").keys()(k))("qty")

                'put raw materials name
            upClause12AYarnConsumptionInformationRangeObject.Range("t" & loopCounter + k).value = "Yarn"

                'put yarn consumption per yds
            upClause12AYarnConsumptionInformationRangeObject.Range("u" & loopCounter + k).value = sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange") _
            (sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange").keys()(k))("yarnConPerYds")
            upClause12AYarnConsumptionInformationRangeObject.Range("u" & loopCounter + k & ":v" & loopCounter + k).Merge

                'put formula for calculate total yarn consumption
            upClause12AYarnConsumptionInformationRangeObject.Range("w" & loopCounter + k).FormulaR1C1 = "=RC[-4]*RC[-2]"
            upClause12AYarnConsumptionInformationRangeObject.Range("w" & loopCounter + k & ":x" & loopCounter + k).Merge

                'put cell value 6%
            upClause12AYarnConsumptionInformationRangeObject.Range("y" & loopCounter + k).value = "6%"


                'put formula for calculate total yarn consumption with 6%
            upClause12AYarnConsumptionInformationRangeObject.Range("z" & loopCounter + k).FormulaR1C1 = "=RC[-3]*6%+RC[-3]"
            upClause12AYarnConsumptionInformationRangeObject.Range("z" & loopCounter + k & ":aa" & loopCounter + k).Merge
            

        Next k

        loopCounter = loopCounter + sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(j))("consumptionRange").Count

    Next j

    upClause12AYarnConsumptionInformationRangeObject.Range("r" & upClause12AYarnConsumptionInformationRangeObject.Rows.Count).FormulaR1C1 = _
    "=SUM(R[-" & upClause12AYarnConsumptionInformationRangeObject.Rows.Count - 1 & "]C:R[-1]C)"

    upClause12AYarnConsumptionInformationRangeObject.Range("s" & upClause12AYarnConsumptionInformationRangeObject.Rows.Count).FormulaR1C1 = _
    "=SUM(R[-" & upClause12AYarnConsumptionInformationRangeObject.Rows.Count - 1 & "]C:R[-1]C)"

    upClause12AYarnConsumptionInformationRangeObject.Range("z" & upClause12AYarnConsumptionInformationRangeObject.Rows.Count).FormulaR1C1 = _
    "=SUM(R[-" & upClause12AYarnConsumptionInformationRangeObject.Rows.Count - 1 & "]C:R[-1]C)"

    Set upClause12AYarnConsumptionInformationRangeObject = upClause12AYarnConsumptionInformationRangeObject.Offset(-1, 1) _
    .Resize(upClause12AYarnConsumptionInformationRangeObject.Rows.Count + 1, upClause12AYarnConsumptionInformationRangeObject.Columns.Count - 1)

    Application.Run "utility_formating_fun.SetBorderInsideHairlineAroundThin", upClause12AYarnConsumptionInformationRangeObject

End Function

Private Function dealWithUpClause12b(ws As Worksheet, sourceDataAsDicUpIssuingStatus As Object)

    Dim upClause12BYarnConsumptionInformationRangeObject As Range
    Set upClause12BYarnConsumptionInformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause12BChemicalDyesConsumptioninformationRangeObjectFromProvidedWs", ws)

    Dim buyerCellsDict As Object
    Set buyerCellsDict = CreateObject("Scripting.Dictionary")

        'take all buyer cells
    buyerCellsDict.Add 0, upClause12BYarnConsumptionInformationRangeObject.Range("b2")
    buyerCellsDict.Add 1, upClause12BYarnConsumptionInformationRangeObject.Range("b8")
    buyerCellsDict.Add 2, upClause12BYarnConsumptionInformationRangeObject.Range("b16")
    buyerCellsDict.Add 3, upClause12BYarnConsumptionInformationRangeObject.Range("b22")
    buyerCellsDict.Add 4, upClause12BYarnConsumptionInformationRangeObject.Range("b31")

    Dim dicKey As Variant

    For Each dicKey In buyerCellsDict.keys
            'clear previous buyer name
        buyerCellsDict(dicKey).value = Null

    Next dicKey

    Dim i As Long

    For i = 0 To sourceDataAsDicUpIssuingStatus.Count - 1
            'put current buyer
        buyerCellsDict(i).value = sourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus.keys()(i))("NameofBuyers")

    Next i

End Function

Private Function dealWithUpClause13(ws As Worksheet, newUpClause8InfoClassifiedPartDic As Object)

    Dim upClause13InformationRangeObject As Range
    Set upClause13InformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause13UseRawMaterialsinformationRangeObjectFromProvidedWs", ws)

        'used in this UP Qty. update
    upClause13InformationRangeObject.Range("o1").value = newUpClause8InfoClassifiedPartDic("yarnImportQty")
    upClause13InformationRangeObject.Range("o2").value = newUpClause8InfoClassifiedPartDic("yarnLocalQty")
    upClause13InformationRangeObject.Range("o3").value = newUpClause8InfoClassifiedPartDic("dyesQty")
    upClause13InformationRangeObject.Range("o4").value = newUpClause8InfoClassifiedPartDic("chemicalsImportQty")
    upClause13InformationRangeObject.Range("o5").value = newUpClause8InfoClassifiedPartDic("chemicalsLocalQty")
    upClause13InformationRangeObject.Range("o6").value = newUpClause8InfoClassifiedPartDic("stretchWrappingFilmQty")

        'used in this UP value update
    upClause13InformationRangeObject.Range("r1").value = newUpClause8InfoClassifiedPartDic("yarnImportValue")
    upClause13InformationRangeObject.Range("r2").value = newUpClause8InfoClassifiedPartDic("yarnLocalValue")
    upClause13InformationRangeObject.Range("r3").value = newUpClause8InfoClassifiedPartDic("dyesValue")
    upClause13InformationRangeObject.Range("r4").value = newUpClause8InfoClassifiedPartDic("chemicalsImportValue")
    upClause13InformationRangeObject.Range("r5").value = newUpClause8InfoClassifiedPartDic("chemicalsLocalValue")
    upClause13InformationRangeObject.Range("r6").value = newUpClause8InfoClassifiedPartDic("stretchWrappingFilmValue")

End Function

Private Function isLastUpUsedUpdatedInImportPerformance(importPerformanceAnyRawMaterialsData As Variant, upNo As Variant, upColumn As Integer) As Boolean
    'return true if only given UP No. updated in import performance used as last UP.
    Dim allExistedUpDict As Object
    Set allExistedUpDict = CreateObject("Scripting.Dictionary")

    Dim i As Long

    For i = 1 To UBound(importPerformanceAnyRawMaterialsData)
            'if UP exist, then add
        If Application.Run("general_utility_functions.isStrPatternExist", importPerformanceAnyRawMaterialsData(i, upColumn), "^" & Replace(upNo, "/", "\/") & "$", True, True, True) Then

            allExistedUpDict(i) = importPerformanceAnyRawMaterialsData(i, upColumn)

        End If

    Next i

    If allExistedUpDict.Count = 0 Then
            'not found
        isLastUpUsedUpdatedInImportPerformance = False
        Exit Function

    End If

    Dim dicKey As Variant

    For Each dicKey In allExistedUpDict.keys

            'if UP exist, but not end of the string, then enter
        If Not Application.Run("general_utility_functions.isStrPatternExist", allExistedUpDict(dicKey), Replace(upNo, "/", "\/") & "$", True, True, False) Then

            isLastUpUsedUpdatedInImportPerformance = False
            Exit Function

        End If

    Next dicKey

        'UP exit only end of the string
    isLastUpUsedUpdatedInImportPerformance = True

End Function