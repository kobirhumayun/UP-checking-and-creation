Attribute VB_Name = "helperFunctionCompareData"
Option Explicit

Private Function upClause6And7CompareWithSource(arrUpClause6Range As Variant, arrUpClause7Range As Variant, sourceData As Variant) As Variant
'      this function give compare result of UP clause 6 & 7 with source data
    Dim arrUpClause6, arrUpClause7 As Variant
    arrUpClause6 = arrUpClause6Range.value
    arrUpClause7 = arrUpClause7Range.value

    Dim regex As New RegExp
    regex.Global = True
    regex.MultiLine = True
    
    
    Dim patternStr As String
    Dim asUpAllLcIndexInSourceData As String
    Dim temp As Variant
    Dim emptyIndex As Variant
    Dim Result As Variant
    
    Dim intialReturnArr() As Variant
    ReDim intialReturnArr(1 To 500, 1 To 4)
    intialReturnArr(1, 1) = "Topic"
    intialReturnArr(1, 2) = "UP Data"
    intialReturnArr(1, 3) = "Source Data"
    intialReturnArr(1, 4) = "Result"
    

    intialReturnArr(2, 1) = " "
    intialReturnArr(2, 2) = "Export LC Information(UP Clause 6 & 7)"
    intialReturnArr(2, 3) = ""
    intialReturnArr(2, 4) = ""
    

    Dim clause7OddFiltered, clause7EvenFiltered As Variant
    clause7OddFiltered = Application.Run("utilityFunction.evenOrOddIndexArrayFilter", arrUpClause7, "odd", False)
    clause7EvenFiltered = Application.Run("utilityFunction.evenOrOddIndexArrayFilter", arrUpClause7, "even", False)
    
    
'   UP Clause7 compare "start"
        'Qty. (start)
        Dim totalExportQtyFromSourceData, totalExportQtyFromUpClause7 As String

        totalExportQtyFromSourceData = Application.Run("utilityFunction.sumQty", sourceData, 9, 27)
        totalExportQtyFromUpClause7 = arrUpClause7(UBound(arrUpClause7, 1), 17)
        
        Result = CLng(totalExportQtyFromSourceData) = CLng(totalExportQtyFromUpClause7)
        
            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & CLng(totalExportQtyFromSourceData) - CLng(totalExportQtyFromUpClause7)
            End If
            
        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range(UBound(arrUpClause7, 1), 17), Result
        
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
        intialReturnArr(emptyIndex, 1) = "Total Export Qty."
        intialReturnArr(emptyIndex, 2) = totalExportQtyFromUpClause7
        intialReturnArr(emptyIndex, 3) = totalExportQtyFromSourceData
        intialReturnArr(emptyIndex, 4) = Result
        'Qty. (end)
        
        'Value (start)
        Dim totalExportValueFromSourceData, totalExportValueFromUpClause7 As String
        totalExportValueFromSourceData = Application.Run("utilityFunction.sumArrColumn", sourceData, 6)
        totalExportValueFromUpClause7 = arrUpClause7(UBound(arrUpClause7, 1), 19)
        
        Result = totalExportValueFromSourceData = totalExportValueFromUpClause7
        
            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & totalExportValueFromSourceData - totalExportValueFromUpClause7
            End If
            
        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range(UBound(arrUpClause7, 1), 19), Result
        
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
        intialReturnArr(emptyIndex, 1) = "Total Export Value"
        intialReturnArr(emptyIndex, 2) = totalExportValueFromUpClause7
        intialReturnArr(emptyIndex, 3) = totalExportValueFromSourceData
        intialReturnArr(emptyIndex, 4) = Result
        'Value (end)
    
    
    
    
    
    Dim i As Integer
    For i = 1 To UBound(clause7OddFiltered, 1) - 1
    
'    LC index no. finding from source data (reverse direction) #start#
    patternStr = ".+"
    regex.pattern = patternStr
    Set temp = regex.Execute(clause7OddFiltered(i, 3))
    
    Dim lcNoFromUpClause7 As String
    lcNoFromUpClause7 = temp.Item(0)
    patternStr = lcNoFromUpClause7
    
    Dim lcIndexInSourceData As Variant
    lcIndexInSourceData = Application.Run("utilityFunction.indexOfReverseOrder", sourceData, Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", patternStr), 4, UBound(sourceData, 1), 1)
     
     
    If IsNull(lcIndexInSourceData) Then
        'if LC not found in source data then this block active
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "LC"
        intialReturnArr(emptyIndex, 2) = i & ") " & lcNoFromUpClause7
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = "Not found in source data"
        
        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("c" & i * 2 - 1), "Mismatch"

        GoTo skipIteration
        
    End If


    asUpAllLcIndexInSourceData = asUpAllLcIndexInSourceData & " " & lcIndexInSourceData
    
    
'    LC index no. finding from source data (reverse direction) #end#
    
    
    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "LC"
    intialReturnArr(emptyIndex, 2) = i & ") " & lcNoFromUpClause7
    intialReturnArr(emptyIndex, 3) = ""
    intialReturnArr(emptyIndex, 4) = ""
    
    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("c" & i * 2 - 1), "OK"
    
    
'    buyer name compare #start#
    Dim buyerNameFromSourceData, buyerNameFromUpClause6 As String
    buyerNameFromSourceData = sourceData(lcIndexInSourceData, 2)
    
    
    If IsArray(arrUpClause6) Then
        buyerNameFromUpClause6 = arrUpClause6(i, 1)
    Else
        buyerNameFromUpClause6 = arrUpClause6
    End If
    
    
    regex.pattern = "^\d\)"
    buyerNameFromUpClause6 = regex.Replace(Trim(buyerNameFromUpClause6), "")
    
    patternStr = buyerNameFromSourceData
    regex.pattern = "^" & Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", patternStr)
    Result = regex.test(Trim(buyerNameFromUpClause6))
    
    
    If Result Then
        Result = "OK"
    Else
        Result = "Mismatch"
    End If
    
    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause6Range.Range("a" & i), Result
    
    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
    intialReturnArr(emptyIndex, 1) = "Buyer Name"
    intialReturnArr(emptyIndex, 2) = buyerNameFromUpClause6
    intialReturnArr(emptyIndex, 3) = buyerNameFromSourceData
    intialReturnArr(emptyIndex, 4) = Result
'    buyer name compare #end#
    
    
    'Bank (start)
    Dim bankNameFromSourceData, bankNameFromUpClause7 As String
    bankNameFromSourceData = sourceData(lcIndexInSourceData, 3)
    bankNameFromUpClause7 = clause7OddFiltered(i, 11)

    patternStr = bankNameFromSourceData
    regex.pattern = "^" & Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", patternStr)
    Result = regex.test(Trim(bankNameFromUpClause7))

    If Result Then
        Result = "OK"
    Else
        Result = "Mismatch"
    End If
    
    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("k" & i * 2 - 1), Result

    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "Bank Name"
    intialReturnArr(emptyIndex, 2) = bankNameFromUpClause7
    intialReturnArr(emptyIndex, 3) = bankNameFromSourceData
    intialReturnArr(emptyIndex, 4) = Result
    'Bank (end)
    
    
    'Shipment Date (start)
    Dim shipmentDtFromSourceData, shipmentDtFromUpClause7 As Date
    shipmentDtFromSourceData = DateValue(sourceData(lcIndexInSourceData, 7))
    shipmentDtFromUpClause7 = DateValue(clause7OddFiltered(i, 15))

    Result = shipmentDtFromSourceData = shipmentDtFromUpClause7

    If Result Then
        Result = "OK"
    Else
        Result = "Mismatch"
    End If
    
    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("o" & i * 2 - 1), Result

    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "Shipment Date"
    intialReturnArr(emptyIndex, 2) = shipmentDtFromUpClause7
    intialReturnArr(emptyIndex, 3) = shipmentDtFromSourceData
    intialReturnArr(emptyIndex, 4) = Result
    'Shipment Date (end)
    
    
    'Expiry Date (start)
    Dim expiryDtFromSourceData, expiryDtFromUpClause7 As Date
    expiryDtFromSourceData = DateValue(sourceData(lcIndexInSourceData, 8))
    expiryDtFromUpClause7 = DateValue(clause7EvenFiltered(i, 15))

    Result = expiryDtFromSourceData = expiryDtFromUpClause7

    If Result Then
        Result = "OK"
    Else
        Result = "Mismatch"
    End If
    
    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("o" & i * 2), Result

    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "Expiry Date"
    intialReturnArr(emptyIndex, 2) = expiryDtFromUpClause7
    intialReturnArr(emptyIndex, 3) = expiryDtFromSourceData
    intialReturnArr(emptyIndex, 4) = Result
    'Expiry Date (end)
    
    
    'Qty. by LC (start)
    Dim filteredLcForQtyFromSourceData, filteredLcForQtyFromUpClause7 As Variant
    filteredLcForQtyFromSourceData = Application.Run("utilityFunction.towDimensionalArrayFilter", sourceData, Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", lcNoFromUpClause7), 4)
    filteredLcForQtyFromUpClause7 = Application.Run("utilityFunction.towDimensionalArrayFilter", clause7OddFiltered, Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", lcNoFromUpClause7), 3)

    Dim qtyByLCFromSourceData, qtyByLCFromUpClause7 As String
    qtyByLCFromSourceData = Application.Run("utilityFunction.sumArrColumn", filteredLcForQtyFromSourceData, 9)
    
    If Not IsEmpty(clause7EvenFiltered(i, 17)) Then
    '   if qty. unit in Mtr then active this code block
       Dim qtyUnitMtr As Variant
        regex.pattern = "[a-z|A-Z|' ']+"
        
        Dim j As Integer
        For j = 1 To UBound(filteredLcForQtyFromUpClause7, 1)
        
          qtyUnitMtr = regex.Replace(filteredLcForQtyFromUpClause7(j, 17), "")
          filteredLcForQtyFromUpClause7(j, 17) = qtyUnitMtr
          
          
        Next j
        
        
            'check mtr to yds qty. converted ok or not start
            Dim mtrQtyConvertedToYdsFromClause7, mtrQtyConvertedToYdsActually As Variant
            
            mtrQtyConvertedToYdsFromClause7 = clause7EvenFiltered(i, 17)
            mtrQtyConvertedToYdsActually = Round(regex.Replace(clause7OddFiltered(i, 17), "") * 1.0936132983)
          
          
            Result = Round(mtrQtyConvertedToYdsFromClause7) = Round(mtrQtyConvertedToYdsActually)
        
            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & Round(mtrQtyConvertedToYdsFromClause7) - Round(mtrQtyConvertedToYdsActually)
            End If
            
            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("q" & i * 2), Result
        
            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
            intialReturnArr(emptyIndex, 1) = "Qty. Mtr to Yds by LC"
            intialReturnArr(emptyIndex, 2) = mtrQtyConvertedToYdsFromClause7 & " (LC SL. " & i & " only)"
            intialReturnArr(emptyIndex, 3) = mtrQtyConvertedToYdsActually & " (LC SL. " & i & " only)"
            intialReturnArr(emptyIndex, 4) = Result
            'check mtr to yds qty. converted ok or not end
        

    End If
    
    
    qtyByLCFromUpClause7 = Application.Run("utilityFunction.sumArrColumn", filteredLcForQtyFromUpClause7, 17)
    
    
    Result = qtyByLCFromSourceData = qtyByLCFromUpClause7

    If Result Then
        Result = "OK"
    Else
        Result = "Mismatch = " & qtyByLCFromSourceData - qtyByLCFromUpClause7
    End If
    
    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("q" & i * 2 - 1), Result

    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "Sum Qty. by LC"
    intialReturnArr(emptyIndex, 2) = qtyByLCFromUpClause7 & " (Sum of Same LC)"
    intialReturnArr(emptyIndex, 3) = qtyByLCFromSourceData & " (Sum of Same LC)"
    intialReturnArr(emptyIndex, 4) = Result
    'Qty. by LC (end)
    
    'Value by LC (start)
    Dim filteredLcForValueFromSourceData, filteredLcForValueFromUpClause7 As Variant
    filteredLcForValueFromSourceData = Application.Run("utilityFunction.towDimensionalArrayFilter", sourceData, Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", lcNoFromUpClause7), 4)
    filteredLcForValueFromUpClause7 = Application.Run("utilityFunction.towDimensionalArrayFilter", clause7OddFiltered, Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", lcNoFromUpClause7), 3)

    Dim valueByLCFromSourceData, valueByLCFromUpClause7 As String
    valueByLCFromSourceData = Application.Run("utilityFunction.sumArrColumn", filteredLcForValueFromSourceData, 6)

    valueByLCFromUpClause7 = Application.Run("utilityFunction.sumArrColumn", filteredLcForValueFromUpClause7, 19)


    Result = valueByLCFromSourceData = valueByLCFromUpClause7

    If Result Then
        Result = "OK"
    Else
        Result = "Mismatch = " & valueByLCFromSourceData - valueByLCFromUpClause7
    End If
    
    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("s" & i * 2 - 1), Result

    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "Sum Value by LC"
    intialReturnArr(emptyIndex, 2) = valueByLCFromUpClause7 & " (Sum of Same LC)"
    intialReturnArr(emptyIndex, 3) = valueByLCFromSourceData & " (Sum of Same LC)"
    intialReturnArr(emptyIndex, 4) = Result
    'Value by LC (end)
    
    
    
    '    M. LC EXP or IP start
    
    Dim ip, exp As Variant
    
    regex.pattern = "IP\:"
    ip = regex.test(sourceData(lcIndexInSourceData, 17))
    
    regex.pattern = "EXP\:"
    exp = regex.test(sourceData(lcIndexInSourceData, 17))
    
    Dim regExReturnedExtractedIp, regExReturnedExtractedExp As Variant
    Dim expReturnArr, expReturnDateStr As Variant
    Dim ipReturnArr, ipReturnDateStr As Variant
        
    If ip Then
        

        regExReturnedExtractedExp = Application.Run("utilityFunction.expOrIpExtractorFromSourceData", sourceData(lcIndexInSourceData, 17), "exp")
        
        
        
        expReturnDateStr = Application.Run("utilityFunction.expOrIpDateExtractorFromSourceDate", sourceData(lcIndexInSourceData, 17), sourceData(lcIndexInSourceData, 18), regExReturnedExtractedExp)
        expReturnArr = Application.Run("utilityFunction.mLcUdExpIpCompareWithSource", clause7OddFiltered(i, 21), regExReturnedExtractedExp, expReturnDateStr, "EXP:")
        intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, expReturnArr, 1)
        
        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("u" & i * 2 - 1), Application.Run("utilityFunction.isAllResultOk", expReturnArr)
        
        
        
        regExReturnedExtractedIp = Application.Run("utilityFunction.expOrIpExtractorFromSourceData", sourceData(lcIndexInSourceData, 17), "ip")
        
        
        
        ipReturnDateStr = Application.Run("utilityFunction.expOrIpDateExtractorFromSourceDate", sourceData(lcIndexInSourceData, 17), sourceData(lcIndexInSourceData, 18), regExReturnedExtractedIp)
        ipReturnArr = Application.Run("utilityFunction.mLcUdExpIpCompareWithSource", clause7OddFiltered(i, 24), regExReturnedExtractedIp, ipReturnDateStr, "IP:")
        intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, ipReturnArr, 1)
        
        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("x" & i * 2 - 1), Application.Run("utilityFunction.isAllResultOk", ipReturnArr)
    
    ElseIf exp Then
    
    
        
        regExReturnedExtractedExp = Application.Run("utilityFunction.expOrIpExtractorFromSourceData", sourceData(lcIndexInSourceData, 17), "exp")
        
        expReturnDateStr = Application.Run("utilityFunction.expOrIpDateExtractorFromSourceDate", sourceData(lcIndexInSourceData, 17), sourceData(lcIndexInSourceData, 18), regExReturnedExtractedExp)
        expReturnArr = Application.Run("utilityFunction.mLcUdExpIpCompareWithSource", clause7OddFiltered(i, 21), regExReturnedExtractedExp, expReturnDateStr, "EXP:")
        intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, expReturnArr, 1)
        
        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("u" & i * 2 - 1), Application.Run("utilityFunction.isAllResultOk", expReturnArr)
    
    
    Else
    
    
        Dim mLcReturnArr As Variant
        mLcReturnArr = Application.Run("utilityFunction.mLcUdExpIpCompareWithSource", clause7OddFiltered(i, 21), sourceData(lcIndexInSourceData, 14), sourceData(lcIndexInSourceData, 15), "Master LC")
        intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, mLcReturnArr, 1)
        
        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause7Range.Range("u" & i * 2 - 1), Application.Run("utilityFunction.isAllResultOk", mLcReturnArr)
        
    End If
    '    M. LC EXP or IP end
    
    
    
    
'   UP Clause7 compare "end"
    
skipIteration:

    Next i



    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
    intialReturnArr(emptyIndex, 1) = "As per UP all LC Sl. in Source"
    intialReturnArr(emptyIndex, 2) = Trim(asUpAllLcIndexInSourceData)
    intialReturnArr(emptyIndex, 3) = ""
    intialReturnArr(emptyIndex, 4) = ""




    Dim intialReturnArrCropIndex As Integer
    intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"
    
    
    upClause6And7CompareWithSource = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)
    

End Function



Private Function upClause8CompareWithSource(arrUpClause8Range As Variant, sourceDataUpIssuingStatus As Variant, sourceDataYarnImport As Variant, sourceDataYarnLocal As Variant, sourceDataDyes As Variant, sourceDataChemicalsImport As Variant, sourceDataChemicalsLocal As Variant, sourceDataStretchWrappingFilm As Variant, sourceDataPreviousUpClause8 As Variant) As Variant
'    this function give compare result of UP clause 8 with source data

    Dim arrUpClause8 As Variant
    arrUpClause8 = arrUpClause8Range.value

    Dim regex As New RegExp
    regex.Global = True
    regex.MultiLine = True


    Dim patternStr As String
    Dim temp As Variant
    Dim emptyIndex As Variant
    Dim Result As Variant
    
    
    Dim intialReturnArr() As Variant
    ReDim intialReturnArr(1 To 1000, 1 To 4)
    intialReturnArr(1, 1) = "Topic"
    intialReturnArr(1, 2) = "UP Data"
    intialReturnArr(1, 3) = "Source Data"
    intialReturnArr(1, 4) = "Result"


    intialReturnArr(2, 1) = " "
    intialReturnArr(2, 2) = "Import LC Information(UP Clause 8)"
    intialReturnArr(2, 3) = ""
    intialReturnArr(2, 4) = ""
    
'    Divide all classifying part start
    Dim upClause8totalYarn, upClause8importYarn, upClause8localYarn, upClause8dyes, upClause8totalChemical, upClause8importChemical, upClause8localChemical, upClause8stretchWrappingFilm As Variant
    Dim upClause8totalYarnSum, upClause8importYarnSum, upClause8localYarnSum, upClause8dyesSum, upClause8totalChemicalSum, upClause8importChemicalSum, upClause8localChemicalSum, upClause8stretchWrappingFilmSum As Variant
    Dim firstIndexTotalYarn, lastIndexTotalYarn, firstIndexImportYarn, lastIndexImportYarn, firstIndexLocalYarn, lastIndexLocalYarn, firstIndexDyes, lastIndexDyes, firstIndexTotalChemical, lastIndexTotalChemical, firstIndexupClause8stretchWrappingFilm, lastIndexupClause8stretchWrappingFilm As Integer
    
    firstIndexTotalYarn = 1
    
'    lastIndexTotalYarn = Application.Run("utilityFunction.indexOf", arrUpClause8, "Dyes", 13, 1, UBound(arrUpClause8, 1)) - 1

    lastIndexTotalYarn = Application.Run("utilityFunction.indexOfReverseOrder", arrUpClause8, "Yarn", 13, UBound(arrUpClause8, 1), 1)
    
'    firstIndexImportYarn = Application.Run("utilityFunction.indexOf", arrUpClause8, "^C-", 6, 1, UBound(arrUpClause8, 1))
    
'    lastIndexImportYarn = lastIndexTotalYarn
    
'    firstIndexLocalYarn = 1
'
'    lastIndexLocalYarn = firstIndexImportYarn - 1


    
'    firstIndexDyes = lastIndexTotalYarn + 1
    
    firstIndexDyes = Application.Run("utilityFunction.indexOf", arrUpClause8, "Dyes", 13, 1, UBound(arrUpClause8, 1))
    
'    lastIndexDyes = Application.Run("utilityFunction.indexOf", arrUpClause8, ".", 13, firstIndexDyes + 1, UBound(arrUpClause8, 1)) - 1
    
    lastIndexDyes = Application.Run("utilityFunction.indexOfReverseOrder", arrUpClause8, "Dyes", 13, UBound(arrUpClause8, 1), 1)
    
    firstIndexTotalChemical = lastIndexDyes + 1
    
    lastIndexTotalChemical = Application.Run("utilityFunction.indexOf", arrUpClause8, "Stretch Wrapping Film", 13, 1, UBound(arrUpClause8, 1)) - 1
    
    firstIndexupClause8stretchWrappingFilm = lastIndexTotalChemical + 1
    
    lastIndexupClause8stretchWrappingFilm = Application.Run("utilityFunction.indexOfReverseOrder", arrUpClause8, "Stretch Wrapping Film", 13, UBound(arrUpClause8, 1), 1)
    
    
    
    upClause8totalYarn = Application.Run("utilityFunction.cropedArryWithStoreLastRow", arrUpClause8, firstIndexTotalYarn, lastIndexTotalYarn)
    
    upClause8importYarn = Application.Run("utilityFunction.towDimensionalArrayFilter", upClause8totalYarn, "^C-", 6)
    
    upClause8localYarn = Application.Run("utilityFunction.towDimensionalArrayFilterNegative", upClause8totalYarn, "^C-", 6)
    
    
    upClause8dyes = Application.Run("utilityFunction.cropedArryWithStoreLastRow", arrUpClause8, firstIndexDyes, lastIndexDyes)
    
    upClause8totalChemical = Application.Run("utilityFunction.cropedArryWithStoreLastRow", arrUpClause8, firstIndexTotalChemical, lastIndexTotalChemical)
    
    upClause8importChemical = Application.Run("utilityFunction.towDimensionalArrayFilter", upClause8totalChemical, "^C-", 6)
    
    upClause8localChemical = Application.Run("utilityFunction.towDimensionalArrayFilter", upClause8totalChemical, "^M-", 6)
    
    upClause8stretchWrappingFilm = Application.Run("utilityFunction.cropedArryWithStoreLastRow", arrUpClause8, firstIndexupClause8stretchWrappingFilm, lastIndexupClause8stretchWrappingFilm)
'    Divide all classifying part end


'    Divided all classifying part compare start
    'check by Qty.
    
    upClause8totalYarnSum = Application.Run("utilityFunction.sumArrColumn", upClause8totalYarn, 15)
    
    upClause8importYarnSum = Application.Run("utilityFunction.sumArrColumn", upClause8importYarn, 15)
    
    If IsArray(upClause8localYarn) Then
    ' error handling if local yarn not exist
        upClause8localYarnSum = Application.Run("utilityFunction.sumArrColumn", upClause8localYarn, 15)
    Else
        upClause8localYarnSum = 0
    End If
    
    upClause8dyesSum = Application.Run("utilityFunction.sumArrColumn", upClause8dyes, 15)
    
    upClause8totalChemicalSum = Application.Run("utilityFunction.sumArrColumn", upClause8totalChemical, 15)
    
    upClause8importChemicalSum = Application.Run("utilityFunction.sumArrColumn", upClause8importChemical, 15)
    
    upClause8localChemicalSum = Application.Run("utilityFunction.sumArrColumn", upClause8localChemical, 15)
    
    upClause8stretchWrappingFilmSum = Application.Run("utilityFunction.sumArrColumn", upClause8stretchWrappingFilm, 15)
    
    Dim sumOfAllPartQty, sumOfAllQty As Variant
    
    sumOfAllPartQty = upClause8importYarnSum + upClause8localYarnSum + upClause8dyesSum + upClause8importChemicalSum + upClause8localChemicalSum + upClause8stretchWrappingFilmSum
    
    sumOfAllQty = Application.Run("utilityFunction.sumArrColumn", arrUpClause8, 15)
    
    Result = CLng(sumOfAllQty) = CLng(sumOfAllPartQty)

            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & CLng(sumOfAllQty) - CLng(sumOfAllPartQty)
            End If

        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("M1:M" & UBound(arrUpClause8, 1) - 1), Result

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Classifying part"
        intialReturnArr(emptyIndex, 2) = sumOfAllQty & " (Sum of all Qty. & Value)"
        intialReturnArr(emptyIndex, 3) = sumOfAllPartQty & " (Sum of all parts Qty. & Value)"
        intialReturnArr(emptyIndex, 4) = Result
    
'    Divided all classifying part compare end

    

'    Total used sum compare start
    

    Dim clause8UsedThisUpQtySum, clause8UsedThisUpValueSum As Variant
  

    clause8UsedThisUpQtySum = Application.Run("utilityFunction.sumArrColumn", arrUpClause8, 21) - arrUpClause8(UBound(arrUpClause8, 1), 21)
    clause8UsedThisUpValueSum = Application.Run("utilityFunction.sumArrColumn", arrUpClause8, 22) - arrUpClause8(UBound(arrUpClause8, 1), 22)
   

    Result = clause8UsedThisUpQtySum = arrUpClause8(UBound(arrUpClause8, 1), 21)

            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & clause8UsedThisUpQtySum - arrUpClause8(UBound(arrUpClause8, 1), 21)
            End If

        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("U" & UBound(arrUpClause8, 1)), Result

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Qty."
        intialReturnArr(emptyIndex, 2) = arrUpClause8(UBound(arrUpClause8, 1), 21)
        intialReturnArr(emptyIndex, 3) = clause8UsedThisUpQtySum
        intialReturnArr(emptyIndex, 4) = Result
        
        
    Result = clause8UsedThisUpValueSum = arrUpClause8(UBound(arrUpClause8, 1), 22)

            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & clause8UsedThisUpValueSum - arrUpClause8(UBound(arrUpClause8, 1), 22)
            End If

        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("V" & UBound(arrUpClause8, 1)), Result

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Value"
        intialReturnArr(emptyIndex, 2) = arrUpClause8(UBound(arrUpClause8, 1), 22)
        intialReturnArr(emptyIndex, 3) = clause8UsedThisUpValueSum
        intialReturnArr(emptyIndex, 4) = Result

'    Total used sum compare end


'    Local yarn B2B LC's total value Qty. compare with UP issuing status start

    Dim upClause8LocalLcSumQty, upClause8LocalLcSumValue As Variant
    Dim sourceDataUpIssuingStatusLocalLcSumQty, sourceDataUpIssuingStatusLocalLcSumValue As Variant
    
    
    If IsArray(upClause8localYarn) Then
    
        upClause8LocalLcSumQty = Application.Run("utilityFunction.sumArrColumn", upClause8localYarn, 15)
        upClause8LocalLcSumValue = Application.Run("utilityFunction.sumArrColumn", upClause8localYarn, 16)
    
    Else
    
        upClause8LocalLcSumQty = 0
        upClause8LocalLcSumValue = 0
          
    End If
    
    
    sourceDataUpIssuingStatusLocalLcSumQty = Application.Run("utilityFunction.sumArrColumn", sourceDataUpIssuingStatus, 23)
    sourceDataUpIssuingStatusLocalLcSumValue = Application.Run("utilityFunction.sumArrColumn", sourceDataUpIssuingStatus, 22)
    
    
    Result = CLng(upClause8LocalLcSumQty) = CLng(sourceDataUpIssuingStatusLocalLcSumQty) 'Qty. Compare

            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & CLng(upClause8LocalLcSumQty) - CLng(sourceDataUpIssuingStatusLocalLcSumQty)
            End If

    If IsArray(upClause8localYarn) Then

        Dim localLcArrIterator As Integer

        For localLcArrIterator = 1 To UBound(upClause8localYarn, 1)
            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("O" & upClause8localYarn(localLcArrIterator, 27)), Result
        Next localLcArrIterator
    
    End If
    
    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "Local B2B LC's total Qty.(yarn)"
    intialReturnArr(emptyIndex, 2) = upClause8LocalLcSumQty
    intialReturnArr(emptyIndex, 3) = sourceDataUpIssuingStatusLocalLcSumQty
    intialReturnArr(emptyIndex, 4) = Result
    
    
    Result = CLng(upClause8LocalLcSumValue) = CLng(sourceDataUpIssuingStatusLocalLcSumValue) 'Value Compare

            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & CLng(upClause8LocalLcSumValue) - CLng(sourceDataUpIssuingStatusLocalLcSumValue)
            End If

    If IsArray(upClause8localYarn) Then
    
        For localLcArrIterator = 1 To UBound(upClause8localYarn, 1)
            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("P" & upClause8localYarn(localLcArrIterator, 27)), Result
        Next localLcArrIterator
    
    End If
    
    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "Local B2B LC's total Value (yarn)"
    intialReturnArr(emptyIndex, 2) = upClause8LocalLcSumValue
    intialReturnArr(emptyIndex, 3) = sourceDataUpIssuingStatusLocalLcSumValue
    intialReturnArr(emptyIndex, 4) = Result
    
'    Local yarn B2B LC's total value Qty. compare with UP issuing status end



'    Total import yarn used sum taken for next clause compare start

    Dim clause8UsedThisUpImportYarnQtySum, clause8UsedThisUpImportYarnValueSum As Variant
    
    If IsArray(upClause8importYarn) Then
    
    clause8UsedThisUpImportYarnQtySum = Application.Run("utilityFunction.sumArrColumn", upClause8importYarn, 21)
    clause8UsedThisUpImportYarnValueSum = Application.Run("utilityFunction.sumArrColumn", upClause8importYarn, 22)
    
    Else
    
    clause8UsedThisUpImportYarnQtySum = 0
    clause8UsedThisUpImportYarnValueSum = 0
    
    End If
    

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Import Yarn Qty."
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpImportYarnQtySum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""
        
        
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Import Yarn Value"
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpImportYarnValueSum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""

'    Total import yarn used sum taken for next clause compare end





'    Total local yarn used sum taken for next clause compare start

    Dim clause8UsedThisUpLocalYarnQtySum, clause8UsedThisUpLocalYarnValueSum As Variant
    
    If IsArray(upClause8localYarn) Then

    clause8UsedThisUpLocalYarnQtySum = Application.Run("utilityFunction.sumArrColumn", upClause8localYarn, 21)
    clause8UsedThisUpLocalYarnValueSum = Application.Run("utilityFunction.sumArrColumn", upClause8localYarn, 22)

    Else

    clause8UsedThisUpLocalYarnQtySum = 0
    clause8UsedThisUpLocalYarnValueSum = 0

    End If

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Local Yarn Qty."
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpLocalYarnQtySum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""
        
        
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Local Yarn Value"
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpLocalYarnValueSum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""

'    Total local yarn used sum taken for next clause compare end





'    Total dyed used sum taken for next clause compare start

    Dim clause8UsedThisUpDyesQtySum, clause8UsedThisUpDyesValueSum As Variant
    
    If IsArray(upClause8dyes) Then

    clause8UsedThisUpDyesQtySum = Application.Run("utilityFunction.sumArrColumn", upClause8dyes, 21)
    clause8UsedThisUpDyesValueSum = Application.Run("utilityFunction.sumArrColumn", upClause8dyes, 22)

    Else

    clause8UsedThisUpDyesQtySum = 0
    clause8UsedThisUpDyesValueSum = 0

    End If

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Dyes Qty."
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpDyesQtySum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""
        
        
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Dyes Value"
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpDyesValueSum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""

'    Total dyed used sum taken for next clause compare end





'    Total import chemical used sum taken for next clause compare start

    Dim clause8UsedThisUpImportChemicalQtySum, clause8UsedThisUpImportChemicalValueSum As Variant
    
    If IsArray(upClause8importChemical) Then

    clause8UsedThisUpImportChemicalQtySum = Application.Run("utilityFunction.sumArrColumn", upClause8importChemical, 21)
    clause8UsedThisUpImportChemicalValueSum = Application.Run("utilityFunction.sumArrColumn", upClause8importChemical, 22)

    Else

    clause8UsedThisUpImportChemicalQtySum = 0
    clause8UsedThisUpImportChemicalValueSum = 0

    End If

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Import Chemical Qty."
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpImportChemicalQtySum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""
        
        
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Import Chemical Value"
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpImportChemicalValueSum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""

'    Total import chemical used sum taken for next clause compare end



'    Total local chemical used sum taken for next clause compare start

    Dim clause8UsedThisUpLocalChemicalQtySum, clause8UsedThisUpLocalChemicalValueSum As Variant
    
    If IsArray(upClause8localChemical) Then

    clause8UsedThisUpLocalChemicalQtySum = Application.Run("utilityFunction.sumArrColumn", upClause8localChemical, 21)
    clause8UsedThisUpLocalChemicalValueSum = Application.Run("utilityFunction.sumArrColumn", upClause8localChemical, 22)

    Else

    clause8UsedThisUpLocalChemicalQtySum = 0
    clause8UsedThisUpLocalChemicalValueSum = 0

    End If

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Local Chemical Qty."
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpLocalChemicalQtySum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""
        
        
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Local Chemical Value"
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpLocalChemicalValueSum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""

'    Total local chemical used sum taken for next clause compare end





'    Total stretch wrapping film used sum taken for next clause compare start

    Dim clause8UsedThisUpStretchWrappingFilmQtySum, clause8UsedThisUpStretchWrappingFilmValueSum As Variant
    
    If IsArray(upClause8stretchWrappingFilm) Then

    clause8UsedThisUpStretchWrappingFilmQtySum = Application.Run("utilityFunction.sumArrColumn", upClause8stretchWrappingFilm, 21)
    clause8UsedThisUpStretchWrappingFilmValueSum = Application.Run("utilityFunction.sumArrColumn", upClause8stretchWrappingFilm, 22)

    Else

    clause8UsedThisUpStretchWrappingFilmQtySum = 0
    clause8UsedThisUpStretchWrappingFilmValueSum = 0

    End If

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Stretch Wrapping Film Qty."
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpStretchWrappingFilmQtySum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""
        
        
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        intialReturnArr(emptyIndex, 1) = "Total Used Stretch Wrapping Film Value"
        intialReturnArr(emptyIndex, 2) = clause8UsedThisUpStretchWrappingFilmValueSum
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""

'    Total stretch wrapping film used sum taken for next clause compare end


    
    
'     Previous & new inserted Bill of Entry or Mushak devided start

        Dim billOfEntryOrMushakPreviousAndNewInsertedBothInCurrentUp As Variant
        Dim billOfEntryOrMushakFromCurrentUpButExistInPreviousUp, billOfEntryOrMushakNewInsertedInCurrentUp As Variant
        
        billOfEntryOrMushakPreviousAndNewInsertedBothInCurrentUp = Application.Run("utilityFunction.findNewOrPreviousBillOfEntryOrMushak", arrUpClause8Range, sourceDataPreviousUpClause8)
        
        billOfEntryOrMushakFromCurrentUpButExistInPreviousUp = billOfEntryOrMushakPreviousAndNewInsertedBothInCurrentUp(1)
        
        billOfEntryOrMushakNewInsertedInCurrentUp = billOfEntryOrMushakPreviousAndNewInsertedBothInCurrentUp(2)
        
'    Previous & new inserted Bill of Entry or Mushak devided end

    
'    Previous balance transfer Qty. compare start
    
        Dim clause8PreviousBalanceInCurrentUpUsedQtySum, clause8PreviousBalanceInCurrentUpStockQtySum, clause8PreviousBalanceInPreviousUpStocQtykSum As Variant


       
       
        If IsArray(billOfEntryOrMushakFromCurrentUpButExistInPreviousUp) Then
            clause8PreviousBalanceInCurrentUpUsedQtySum = Application.Run("utilityFunction.sumArrColumn", billOfEntryOrMushakFromCurrentUpButExistInPreviousUp, 21)
        Else
            clause8PreviousBalanceInCurrentUpUsedQtySum = 0
        End If
    
       
        If IsArray(billOfEntryOrMushakFromCurrentUpButExistInPreviousUp) Then
            clause8PreviousBalanceInCurrentUpStockQtySum = Application.Run("utilityFunction.sumArrColumn", billOfEntryOrMushakFromCurrentUpButExistInPreviousUp, 25)
        Else
            clause8PreviousBalanceInCurrentUpStockQtySum = 0
        End If
    
       clause8PreviousBalanceInPreviousUpStocQtykSum = Application.Run("utilityFunction.sumArrColumn", sourceDataPreviousUpClause8, 25)
        
    
        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round((Round(clause8PreviousBalanceInCurrentUpUsedQtySum, 2) + Round(clause8PreviousBalanceInCurrentUpStockQtySum, 2)), 2), Round(clause8PreviousBalanceInPreviousUpStocQtykSum, 2), 0.1)
'       result = Round((Round(clause8PreviousBalanceInCurrentUpUsedQtySum, 2) + Round(clause8PreviousBalanceInCurrentUpStockQtySum, 2)), 2) = Round(clause8PreviousBalanceInPreviousUpStocQtykSum, 2)
                
               If Result Then
                   Result = "OK"
               Else
                   Result = "Mismatch = " & Round((Round(clause8PreviousBalanceInCurrentUpUsedQtySum, 2) + Round(clause8PreviousBalanceInCurrentUpStockQtySum, 2)) - Round(clause8PreviousBalanceInPreviousUpStocQtykSum, 2), 2)
               End If
    
        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("Y1:Y" & UBound(arrUpClause8, 1) - 1), Result
    
       emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
       intialReturnArr(emptyIndex, 1) = "Previous Balance Qty. Transfer"
       intialReturnArr(emptyIndex, 2) = Round(clause8PreviousBalanceInCurrentUpUsedQtySum + clause8PreviousBalanceInCurrentUpStockQtySum, 2)
       intialReturnArr(emptyIndex, 3) = Round(clause8PreviousBalanceInPreviousUpStocQtykSum, 2)
       intialReturnArr(emptyIndex, 4) = Result
       
'    Previous balance transfer Qty. compare end

    
'    Previous balance transfer value compare start
    
         Dim clause8PreviousBalanceInCurrentUpUsedValueSum, clause8PreviousBalanceInCurrentUpStockValueSum, clause8PreviousBalanceInPreviousUpStocValuekSum As Variant


           
            If IsArray(billOfEntryOrMushakFromCurrentUpButExistInPreviousUp) Then
                clause8PreviousBalanceInCurrentUpUsedValueSum = Application.Run("utilityFunction.sumArrColumn", billOfEntryOrMushakFromCurrentUpButExistInPreviousUp, 22)
            Else
                clause8PreviousBalanceInCurrentUpUsedValueSum = 0
            End If
            
           
            If IsArray(billOfEntryOrMushakFromCurrentUpButExistInPreviousUp) Then
                clause8PreviousBalanceInCurrentUpStockValueSum = Application.Run("utilityFunction.sumArrColumn", billOfEntryOrMushakFromCurrentUpButExistInPreviousUp, 26)
            Else
                clause8PreviousBalanceInCurrentUpStockValueSum = 0
            End If
    
           clause8PreviousBalanceInPreviousUpStocValuekSum = Application.Run("utilityFunction.sumArrColumn", sourceDataPreviousUpClause8, 26)
        
            Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round((Round(clause8PreviousBalanceInCurrentUpUsedValueSum, 2) + Round(clause8PreviousBalanceInCurrentUpStockValueSum, 2)), 2), Round(clause8PreviousBalanceInPreviousUpStocValuekSum, 2), 0.1)
        
'           result = Round((Round(clause8PreviousBalanceInCurrentUpUsedValueSum, 2) + Round(clause8PreviousBalanceInCurrentUpStockValueSum, 2)), 2) = Round(clause8PreviousBalanceInPreviousUpStocValuekSum, 2)
        
                   If Result Then
                       Result = "OK"
                   Else
                       Result = "Mismatch = " & Round((Round(clause8PreviousBalanceInCurrentUpUsedValueSum, 2) + Round(clause8PreviousBalanceInCurrentUpStockValueSum, 2)) - Round(clause8PreviousBalanceInPreviousUpStocValuekSum, 2), 2)
                   End If
        
            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("Z1:Z" & UBound(arrUpClause8, 1) - 1), Result
        
           emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
           intialReturnArr(emptyIndex, 1) = "Previous Balance Value Transfer"
           intialReturnArr(emptyIndex, 2) = Round(clause8PreviousBalanceInCurrentUpUsedValueSum + clause8PreviousBalanceInCurrentUpStockValueSum, 2)
           intialReturnArr(emptyIndex, 3) = Round(clause8PreviousBalanceInPreviousUpStocValuekSum, 2)
           intialReturnArr(emptyIndex, 4) = Result
'    Previous balance transfer value compare end
    
    
'    New inserted Bill of Entry or Mushak previous used Qty. check start

            If IsArray(billOfEntryOrMushakNewInsertedInCurrentUp) Then
            
            Dim newInsertedIteratorQty As Integer
            Dim newInsertedPreviousUsedQty As Variant
            
            For newInsertedIteratorQty = 1 To UBound(billOfEntryOrMushakNewInsertedInCurrentUp, 1)
                
                newInsertedPreviousUsedQty = billOfEntryOrMushakNewInsertedInCurrentUp(newInsertedIteratorQty, 17)
                
                Result = Round(newInsertedPreviousUsedQty) = 0
                
                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(newInsertedPreviousUsedQty - 0, 2)
                End If
                
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("Q" & billOfEntryOrMushakNewInsertedInCurrentUp(newInsertedIteratorQty, UBound(billOfEntryOrMushakNewInsertedInCurrentUp, 2))), Result
                Application.Run "EditComment", arrUpClause8Range.Range("Q" & billOfEntryOrMushakNewInsertedInCurrentUp(newInsertedIteratorQty, UBound(billOfEntryOrMushakNewInsertedInCurrentUp, 2))), "New Entry Previous Used Qty. " & Result
                
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1))        ' find empty string pattern = "^$"
                
                intialReturnArr(emptyIndex, 1) = "New Entry " & billOfEntryOrMushakNewInsertedInCurrentUp(newInsertedIteratorQty, 6)
                intialReturnArr(emptyIndex, 2) = " "
                intialReturnArr(emptyIndex, 3) = "Previous Used Qty. = " & Round(newInsertedPreviousUsedQty, 2)
                intialReturnArr(emptyIndex, 4) = Result
                
            Next newInsertedIteratorQty
            
            End If

'    New inserted Bill of Entry or Mushak previous used Qty. check end



'    New inserted Bill of Entry or Mushak previous used value check start
            
            If IsArray(billOfEntryOrMushakNewInsertedInCurrentUp) Then
            Dim newInsertedIteratorValue As Integer
            Dim newInsertedPreviousUsedValue As Variant
            
            For newInsertedIteratorValue = 1 To UBound(billOfEntryOrMushakNewInsertedInCurrentUp, 1)
                
                newInsertedPreviousUsedValue = billOfEntryOrMushakNewInsertedInCurrentUp(newInsertedIteratorValue, 18)
                
                Result = Round(newInsertedPreviousUsedValue) = 0
                
                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(newInsertedPreviousUsedValue - 0, 2)
                End If
                
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("R" & billOfEntryOrMushakNewInsertedInCurrentUp(newInsertedIteratorValue, UBound(billOfEntryOrMushakNewInsertedInCurrentUp, 2))), Result
                Application.Run "EditComment", arrUpClause8Range.Range("R" & billOfEntryOrMushakNewInsertedInCurrentUp(newInsertedIteratorValue, UBound(billOfEntryOrMushakNewInsertedInCurrentUp, 2))), "New Entry Previous Used Value " & Result
                
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1))        ' find empty string pattern = "^$"
                
                intialReturnArr(emptyIndex, 1) = "New Entry " & billOfEntryOrMushakNewInsertedInCurrentUp(newInsertedIteratorValue, 6)
                intialReturnArr(emptyIndex, 2) = " "
                intialReturnArr(emptyIndex, 3) = "Previous Used Value = " & Round(newInsertedPreviousUsedValue, 2)
                intialReturnArr(emptyIndex, 4) = Result
                
            Next newInsertedIteratorValue

            End If
            
'    New inserted Bill of Entry or Mushak previous used value check end
 
  
    
'    Transfered & excluded (Take from previous UP), Bill of Entry or Mushak devided start

    Dim billOfEntryOrMushakTransferedAndExcludeBothInPreviousUp As Variant
    Dim billOfEntryOrMushakTransferedFromPreviousUp, billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp As Variant
    
    billOfEntryOrMushakTransferedAndExcludeBothInPreviousUp = Application.Run("utilityFunction.findbillOfEntryOrMushakExcludeOrTransferedFromPreviousUp", arrUpClause8Range, sourceDataPreviousUpClause8)
    
    billOfEntryOrMushakTransferedFromPreviousUp = billOfEntryOrMushakTransferedAndExcludeBothInPreviousUp(1)
    
    billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp = billOfEntryOrMushakTransferedAndExcludeBothInPreviousUp(2)
    
'    Transfered & excluded (Take from previous UP), Bill of Entry or Mushak devided start
    
    
    
    
'    Excluded Bill of Entry or Mushak last balance Qty. check start

            If IsArray(billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp) Then
            
            Dim excludedIteratorQty As Integer
            Dim excludedLastBalanceQty As Variant
            
            For excludedIteratorQty = 1 To UBound(billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp, 1)
                
                excludedLastBalanceQty = billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp(excludedIteratorQty, 25)
                
                Result = Round(excludedLastBalanceQty) = 0
                
                If Result Then
                    Result = "OK"
                Else
                    Result = Round(excludedLastBalanceQty - 0, 2)
                End If
                

                
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1))        ' find empty string pattern = "^$"
                
                intialReturnArr(emptyIndex, 1) = "Excluded " & billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp(excludedIteratorQty, 6)
                intialReturnArr(emptyIndex, 2) = " "
                intialReturnArr(emptyIndex, 3) = "Last Balance Qty. = " & Round(excludedLastBalanceQty, 2)
                intialReturnArr(emptyIndex, 4) = Result
                
            Next excludedIteratorQty
            
            End If


'    Excluded Bill of Entry or Mushak last balance Qty. check end



'    Excluded Bill of Entry or Mushak last balance value check start

            If IsArray(billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp) Then
            
            Dim excludedIteratorValue As Integer
            Dim excludedLastBalanceValue As Variant
            
            For excludedIteratorValue = 1 To UBound(billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp, 1)
                
                excludedLastBalanceValue = billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp(excludedIteratorValue, 26)
                
                Result = Round(excludedLastBalanceValue) = 0
                
                If Result Then
                    Result = "OK"
                Else
                    Result = Round(excludedLastBalanceValue - 0, 2)
                End If
                

                
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1))        ' find empty string pattern = "^$"
                
                intialReturnArr(emptyIndex, 1) = "Excluded " & billOfEntryOrMushakExcludeFromCurrentUpButExistInPreviousUp(excludedIteratorValue, 6)
                intialReturnArr(emptyIndex, 2) = " "
                intialReturnArr(emptyIndex, 3) = "Last Balance Value = " & Round(excludedLastBalanceValue, 2)
                intialReturnArr(emptyIndex, 4) = Result
                
            Next excludedIteratorValue
            
            End If


'    Excluded Bill of Entry or Mushak last balance value check end
    
    
    


'    Previous balance check by Specific Bill of Entry or Mushak start


        Dim previousBalanceTransferReturnArr As Variant

        previousBalanceTransferReturnArr = Application.Run("utilityFunction.upClause8SpecificMushakOrBillOfEntryPreviousBalanceTransferCompare", arrUpClause8Range, sourceDataPreviousUpClause8)
        intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, previousBalanceTransferReturnArr, 1)


'    Previous balance check by Specific Bill of Entry or Mushak end






'    Local yarn B2B LC's value Qty. & date compare with UP issuing status by LC start
Dim b2bLcIterator As Integer
        
For b2bLcIterator = 1 To UBound(sourceDataUpIssuingStatus, 1)



If IsEmpty(sourceDataUpIssuingStatus(b2bLcIterator, 21)) Then
    GoTo skipIteration
End If


Dim b2bLcFromSourceDataUpIssuingStatus As Variant

patternStr = "\d.+"
regex.pattern = patternStr
Set temp = regex.Execute(sourceDataUpIssuingStatus(b2bLcIterator, 20))

If temp.Count = 0 Then
    GoTo skipIteration
End If


b2bLcFromSourceDataUpIssuingStatus = temp.Item(0)


emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

intialReturnArr(emptyIndex, 1) = "B2B LC (yarn)"
intialReturnArr(emptyIndex, 2) = b2bLcIterator & ") " & b2bLcFromSourceDataUpIssuingStatus
intialReturnArr(emptyIndex, 3) = ""
intialReturnArr(emptyIndex, 4) = ""

If IsArray(upClause8localYarn) Then

Dim filteredB2bLcFromUpClause8 As Variant

    filteredB2bLcFromUpClause8 = Application.Run("utilityFunction.towDimensionalArrayFilter", upClause8localYarn, b2bLcFromSourceDataUpIssuingStatus, 1)

Else

    filteredB2bLcFromUpClause8 = Null
    
End If


If IsNull(filteredB2bLcFromUpClause8) Then
    'if LC not found in UP clause8 then this block active
    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "B2B LC"
    intialReturnArr(emptyIndex, 2) = b2bLcIterator & ") " & b2bLcFromSourceDataUpIssuingStatus
    intialReturnArr(emptyIndex, 3) = ""
    intialReturnArr(emptyIndex, 4) = "Not found in UP clause8"
    

    GoTo skipIteration
    
End If


Dim upClause8LocalLcSumQtyByLc, upClause8LocalLcSumValueByLc As Variant
Dim sourceDataUpIssuingStatusLocalLcSumQtyByLc, sourceDataUpIssuingStatusLocalLcSumValueByLc As Variant



upClause8LocalLcSumQtyByLc = Application.Run("utilityFunction.sumArrColumn", filteredB2bLcFromUpClause8, 15)
upClause8LocalLcSumValueByLc = Application.Run("utilityFunction.sumArrColumn", filteredB2bLcFromUpClause8, 16)

sourceDataUpIssuingStatusLocalLcSumQtyByLc = sourceDataUpIssuingStatus(b2bLcIterator, 23)
sourceDataUpIssuingStatusLocalLcSumValueByLc = sourceDataUpIssuingStatus(b2bLcIterator, 22)


Result = Round(upClause8LocalLcSumQtyByLc, 2) = Round(sourceDataUpIssuingStatusLocalLcSumQtyByLc, 2) 'Qty. Compare


If Result Then
    Result = "OK"
Else
    Result = "Mismatch = " & Round(upClause8LocalLcSumQtyByLc, 2) - Round(sourceDataUpIssuingStatusLocalLcSumQtyByLc, 2)
End If

Dim upClause8LocalLcSumQtyByLcResult As Variant

upClause8LocalLcSumQtyByLcResult = Result

emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

intialReturnArr(emptyIndex, 1) = "Qty."
intialReturnArr(emptyIndex, 2) = upClause8LocalLcSumQtyByLc
intialReturnArr(emptyIndex, 3) = sourceDataUpIssuingStatusLocalLcSumQtyByLc
intialReturnArr(emptyIndex, 4) = Result

Result = Round(upClause8LocalLcSumValueByLc, 2) = Round(sourceDataUpIssuingStatusLocalLcSumValueByLc, 2) 'Value Compare



        If Result Then
            Result = "OK"
        Else
            Result = "Mismatch = " & Round(upClause8LocalLcSumValueByLc, 2) - Round(sourceDataUpIssuingStatusLocalLcSumValueByLc, 2)
        End If
     
Dim upClause8LocalLcSumValueByLcResult As Variant

upClause8LocalLcSumValueByLcResult = Result
        
emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

intialReturnArr(emptyIndex, 1) = "Value"
intialReturnArr(emptyIndex, 2) = upClause8LocalLcSumValueByLc
intialReturnArr(emptyIndex, 3) = sourceDataUpIssuingStatusLocalLcSumValueByLc
intialReturnArr(emptyIndex, 4) = Result



Dim localLcArrIteratorByLc As Integer
Dim isAllLcDateOk As Boolean
    isAllLcDateOk = True
For localLcArrIteratorByLc = 1 To UBound(filteredB2bLcFromUpClause8, 1)


    Dim b2bLcDateFromSourceData As Variant

    b2bLcDateFromSourceData = sourceDataUpIssuingStatus(b2bLcIterator, 21)

    patternStr = "^" & Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", b2bLcDateFromSourceData) & "$"

    regex.pattern = patternStr
    
    Result = regex.test(filteredB2bLcFromUpClause8(localLcArrIteratorByLc, 1))

    If Not Result Then

        isAllLcDateOk = False
    
    End If

    Dim upClause8LocalLcDateResult As Variant

    If Result Then
        upClause8LocalLcDateResult = "OK"
    Else
        upClause8LocalLcDateResult = "Mismatch"
    End If



    Result = upClause8LocalLcSumQtyByLcResult = "OK" And upClause8LocalLcSumValueByLcResult = "OK" And upClause8LocalLcDateResult = "OK"

    If Result Then
        Result = "OK"
    Else
        Result = "Mismatch"
    End If

    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("A" & filteredB2bLcFromUpClause8(localLcArrIteratorByLc, UBound(filteredB2bLcFromUpClause8, 2))), Result
    
    
    Application.Run "EditComment", arrUpClause8Range.Range("A" & filteredB2bLcFromUpClause8(localLcArrIteratorByLc, UBound(filteredB2bLcFromUpClause8, 2))), "B2B LC Qty. Sum by LC " & upClause8LocalLcSumQtyByLcResult
    Application.Run "EditComment", arrUpClause8Range.Range("A" & filteredB2bLcFromUpClause8(localLcArrIteratorByLc, UBound(filteredB2bLcFromUpClause8, 2))), "B2B LC Value Sum by LC " & upClause8LocalLcSumValueByLcResult
    Application.Run "EditComment", arrUpClause8Range.Range("A" & filteredB2bLcFromUpClause8(localLcArrIteratorByLc, UBound(filteredB2bLcFromUpClause8, 2))), "B2B LC Date " & upClause8LocalLcDateResult
    
Next localLcArrIteratorByLc

    Result = isAllLcDateOk

    If Result Then
        Result = "OK"
    Else
        Result = "Mismatch"
    End If


    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

    intialReturnArr(emptyIndex, 1) = "B2B LC Date"
    intialReturnArr(emptyIndex, 2) = " "
    intialReturnArr(emptyIndex, 3) = " "
    intialReturnArr(emptyIndex, 4) = Result

    isAllLcDateOk = True ' reset

skipIteration:

Next b2bLcIterator

'    Local yarn B2B LC's value Qty. & date compare with UP issuing status by LC end






'   Local yarn mushak information compare whith import performance start

    If IsArray(upClause8localYarn) Then
        ' error handling if local yarn not exist
        Dim intialReturnArrAfterLocalYarn As Variant
        intialReturnArrAfterLocalYarn = Application.Run("utilityFunction.upClause8MushakOrBillOfEntryCompare", arrUpClause8Range, upClause8localYarn, sourceDataYarnLocal, 3, 4, 7, 8, "Mushak", "Yarn")
        intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, intialReturnArrAfterLocalYarn, 1)
    End If
    
'   Local yarn mushak information compare whith import performance end




'   Import yarn bill of entry information compare whith import performance start


    Dim intialReturnArrAfterImportYarn As Variant
    intialReturnArrAfterImportYarn = Application.Run("utilityFunction.upClause8MushakOrBillOfEntryCompare", arrUpClause8Range, upClause8importYarn, sourceDataYarnImport, 3, 4, 7, 8, "Bill of Entry", "Yarn")
    intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, intialReturnArrAfterImportYarn, 1)

'   Import yarn bill of entry information compare whith import performance end




'   Dyes bill of entry information compare whith import performance start


    Dim intialReturnArrAfterDyes As Variant
    intialReturnArrAfterDyes = Application.Run("utilityFunction.upClause8MushakOrBillOfEntryCompare", arrUpClause8Range, upClause8dyes, sourceDataDyes, 3, 4, 7, 8, "Bill of Entry", "Dyes")
    intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, intialReturnArrAfterDyes, 1)

'   Dyes bill of entry information compare whith import performance end






'   Chemical bill of entry information compare whith import performance start


    Dim intialReturnArrAfterImportChemical As Variant
    intialReturnArrAfterImportChemical = Application.Run("utilityFunction.upClause8MushakOrBillOfEntryCompare", arrUpClause8Range, upClause8importChemical, sourceDataChemicalsImport, 3, 4, 8, 9, "Bill of Entry", "Chemical")
    intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, intialReturnArrAfterImportChemical, 1)

'   Chemical bill of entry information compare whith import performance end






'   Chemical local mushak information compare whith import performance start


    Dim intialReturnArrAfterLocalChemical As Variant
    intialReturnArrAfterLocalChemical = Application.Run("utilityFunction.upClause8MushakOrBillOfEntryCompare", arrUpClause8Range, upClause8localChemical, sourceDataChemicalsLocal, 3, 4, 8, 9, "Mushak", "Chemical")
    intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, intialReturnArrAfterLocalChemical, 1)

'   Chemical local mushak information compare whith import performance end



'   Stretch Wrapping Film bill of entry information compare whith import performance start


Dim intialReturnArrAfterStretchWrappingFilm As Variant
intialReturnArrAfterStretchWrappingFilm = Application.Run("utilityFunction.upClause8MushakOrBillOfEntryCompare", arrUpClause8Range, upClause8stretchWrappingFilm, sourceDataStretchWrappingFilm, 3, 4, 8, 9, "Bill of Entry", "Stretch Wrapping Film")
intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, intialReturnArrAfterStretchWrappingFilm, 1)

'   Stretch Wrapping Film bill of entry information compare whith import performance end


    
Dim intialReturnArrCropIndex As Integer
intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"


upClause8CompareWithSource = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)

End Function






Private Function upClause9CompareWithSource(arrUpClause9Range As Variant, resultClause8 As Variant, upYarnConsumptionInformation As Variant, sourceDataPreviousUpClause9 As Variant, sourceDataImportPerformanceTotalSummary As Variant) As Variant
    '    this function give compare result of UP clause 9 with source data

        Dim arrUpClause9 As Variant
        arrUpClause9 = arrUpClause9Range.value

        Dim regex As New RegExp
        regex.Global = True
        regex.MultiLine = True


        Dim emptyIndex As Variant
        Dim Result As Variant

        Dim intialReturnArr(1 To 50, 1 To 4) As Variant
        intialReturnArr(1, 1) = "Topic"
        intialReturnArr(1, 2) = "UP Data"
        intialReturnArr(1, 3) = "Source Data"
        intialReturnArr(1, 4) = "Result"


        intialReturnArr(2, 1) = " "
        intialReturnArr(2, 2) = "Stock Information(UP Clause 9)"
        intialReturnArr(2, 3) = ""
        intialReturnArr(2, 4) = ""



    '    used raw metarial compare with up clause 8 start

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = " "
            intialReturnArr(emptyIndex, 2) = "Used This UP"
            intialReturnArr(emptyIndex, 3) = ""
            intialReturnArr(emptyIndex, 4) = ""

        Dim importYarnUpClause8Qty, localYarnUpClause8Qty, dyesUpClause8Qty, importChemicalUpClause8Qty, localChemicalUpClause8Qty, stretchWrappingFilmUpClause8Qty, totalSumUsedUpClause8 As Variant
        Dim importYarnUpClause9Qty, localYarnUpClause9Qty, dyesUpClause9Qty, importChemicalUpClause9Qty, localChemicalUpClause9Qty, stretchWrappingFilmUpClause9Qty, totalSumUsedUpClause9 As Variant

        importYarnUpClause8Qty = resultClause8(8, 2)
        localYarnUpClause8Qty = resultClause8(10, 2)
        dyesUpClause8Qty = resultClause8(12, 2)
        importChemicalUpClause8Qty = resultClause8(14, 2)
        localChemicalUpClause8Qty = resultClause8(16, 2)
        stretchWrappingFilmUpClause8Qty = resultClause8(18, 2)
        totalSumUsedUpClause8 = importYarnUpClause8Qty + localYarnUpClause8Qty + dyesUpClause8Qty + importChemicalUpClause8Qty + localChemicalUpClause8Qty + stretchWrappingFilmUpClause8Qty

        importYarnUpClause9Qty = arrUpClause9(1, 23)
        localYarnUpClause9Qty = arrUpClause9(2, 23)
        dyesUpClause9Qty = arrUpClause9(3, 23)
        importChemicalUpClause9Qty = arrUpClause9(4, 23)
        localChemicalUpClause9Qty = arrUpClause9(5, 23)
        stretchWrappingFilmUpClause9Qty = arrUpClause9(6, 23)
        totalSumUsedUpClause9 = arrUpClause9(7, 23)



        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(importYarnUpClause8Qty, 2), Round(importYarnUpClause9Qty, 2), 0.1)

                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(Round(importYarnUpClause8Qty, 2) - Round(importYarnUpClause9Qty, 2), 2)
                End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("W" & 1), Result
            Application.Run "EditComment", arrUpClause9Range.Range("W" & 1), "Checked with UP clause 8 import yarn " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "Import Yarn Used Qty."
            intialReturnArr(emptyIndex, 2) = importYarnUpClause9Qty
            intialReturnArr(emptyIndex, 3) = importYarnUpClause8Qty
            intialReturnArr(emptyIndex, 4) = Result

        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(localYarnUpClause8Qty, 2), Round(localYarnUpClause9Qty, 2), 0.1)

                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(Round(localYarnUpClause8Qty, 2) - Round(localYarnUpClause9Qty, 2), 2)
                End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("W" & 2), Result
            Application.Run "EditComment", arrUpClause9Range.Range("W" & 2), "Checked with UP clause 8 local yarn " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "Local Yarn Used Qty."
            intialReturnArr(emptyIndex, 2) = localYarnUpClause9Qty
            intialReturnArr(emptyIndex, 3) = localYarnUpClause8Qty
            intialReturnArr(emptyIndex, 4) = Result


        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(dyesUpClause8Qty, 2), Round(dyesUpClause9Qty, 2), 0.1)

                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(Round(dyesUpClause8Qty, 2) - Round(dyesUpClause9Qty, 2), 2)
                End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("W" & 3), Result
            Application.Run "EditComment", arrUpClause9Range.Range("W" & 3), "Checked with UP clause 8 Dyes " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "Dyes Used Qty."
            intialReturnArr(emptyIndex, 2) = dyesUpClause9Qty
            intialReturnArr(emptyIndex, 3) = dyesUpClause8Qty
            intialReturnArr(emptyIndex, 4) = Result


        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(importChemicalUpClause8Qty, 2), Round(importChemicalUpClause9Qty, 2), 0.1)

               If Result Then
                   Result = "OK"
               Else
                   Result = "Mismatch = " & Round(Round(importChemicalUpClause8Qty, 2) - Round(importChemicalUpClause9Qty, 2), 2)
               End If

           Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("W" & 4), Result
           Application.Run "EditComment", arrUpClause9Range.Range("W" & 4), "Checked with UP clause 8 import chemical " & Result

           emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

           intialReturnArr(emptyIndex, 1) = "Import Chemical Used Qty."
           intialReturnArr(emptyIndex, 2) = importChemicalUpClause9Qty
           intialReturnArr(emptyIndex, 3) = importChemicalUpClause8Qty
           intialReturnArr(emptyIndex, 4) = Result


        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(localChemicalUpClause8Qty, 2), Round(localChemicalUpClause9Qty, 2), 0.1)

               If Result Then
                   Result = "OK"
               Else
                   Result = "Mismatch = " & Round(Round(localChemicalUpClause8Qty, 2) - Round(localChemicalUpClause9Qty, 2), 2)
               End If

           Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("W" & 5), Result
           Application.Run "EditComment", arrUpClause9Range.Range("W" & 5), "Checked with UP clause 8 local chemical " & Result

           emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

           intialReturnArr(emptyIndex, 1) = "Local Chemical Used Qty."
           intialReturnArr(emptyIndex, 2) = localChemicalUpClause9Qty
           intialReturnArr(emptyIndex, 3) = localChemicalUpClause8Qty
           intialReturnArr(emptyIndex, 4) = Result


        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(stretchWrappingFilmUpClause8Qty, 2), Round(stretchWrappingFilmUpClause9Qty, 2), 0.1)

           If Result Then
               Result = "OK"
           Else
               Result = "Mismatch = " & Round(Round(stretchWrappingFilmUpClause8Qty, 2) - Round(stretchWrappingFilmUpClause9Qty, 2), 2)
           End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("W" & 6), Result
            Application.Run "EditComment", arrUpClause9Range.Range("W" & 6), "Checked with UP clause 8 stretch wrapping film " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "Stretch Wrapping Film Used Qty."
            intialReturnArr(emptyIndex, 2) = stretchWrappingFilmUpClause9Qty
            intialReturnArr(emptyIndex, 3) = stretchWrappingFilmUpClause8Qty
            intialReturnArr(emptyIndex, 4) = Result


        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(totalSumUsedUpClause8, 2), Round(totalSumUsedUpClause9, 2), 0.1)

                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(Round(totalSumUsedUpClause8, 2) - Round(totalSumUsedUpClause9, 2), 2)
                End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("W" & 7), Result
            Application.Run "EditComment", arrUpClause9Range.Range("W" & 7), "Checked with UP clause 8 Qty. sum " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "Total Used Qty."
            intialReturnArr(emptyIndex, 2) = totalSumUsedUpClause9
            intialReturnArr(emptyIndex, 3) = totalSumUsedUpClause8
            intialReturnArr(emptyIndex, 4) = Result



    '    used raw metarial compare with up clause 8 end


    '    total yarn compare with comsumption sheet start
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = " "
            intialReturnArr(emptyIndex, 2) = "Total Yarn Compare With Consumption Sheet"
            intialReturnArr(emptyIndex, 3) = ""
            intialReturnArr(emptyIndex, 4) = ""

        Dim totalYarnClause9, totalYarnConsumption As Variant


        totalYarnClause9 = arrUpClause9(1, 23) + arrUpClause9(2, 23)

        totalYarnConsumption = upYarnConsumptionInformation(1, 8)

        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(totalYarnClause9, 2), Round(totalYarnConsumption, 2), 0.1)

                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(Round(totalYarnClause9, 2) - Round(totalYarnConsumption, 2), 2)
                End If

            Application.Run "EditComment", arrUpClause9Range.Range("W" & 1), "Import & local yarn sum and checked with yarn comsumption sheet " & Result
            Application.Run "EditComment", arrUpClause9Range.Range("W" & 2), "Import & local yarn sum and checked with yarn comsumption sheet " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "Yarn Consumption"
            intialReturnArr(emptyIndex, 2) = totalYarnClause9
            intialReturnArr(emptyIndex, 3) = totalYarnConsumption
            intialReturnArr(emptyIndex, 4) = Result

    '    total yarn compare with comsumption sheet end




    '    previous used raw metarial compare with previous UP start


        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = " "
            intialReturnArr(emptyIndex, 2) = "Previous Used"
            intialReturnArr(emptyIndex, 3) = ""
            intialReturnArr(emptyIndex, 4) = ""


        Dim currentUpPreviousUsed, currentUpPreviousUsedSum, previousUpPreviousUsed, previousUpPreviousUsedSum As Variant

        Dim previousUsedIterator As Integer

        currentUpPreviousUsedSum = 0

        For previousUsedIterator = 1 To UBound(arrUpClause9, 1) - 1

        currentUpPreviousUsed = arrUpClause9(previousUsedIterator, 19)

        currentUpPreviousUsedSum = currentUpPreviousUsedSum + currentUpPreviousUsed

        previousUpPreviousUsed = sourceDataPreviousUpClause9(previousUsedIterator, 27)

        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(currentUpPreviousUsed, 2), Round(previousUpPreviousUsed, 2), 0.1)

                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(Round(currentUpPreviousUsed, 2) - Round(previousUpPreviousUsed, 2), 2)
                End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("S" & previousUsedIterator), Result
            Application.Run "EditComment", arrUpClause9Range.Range("S" & previousUsedIterator), "In Previous UP Used " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = arrUpClause9(previousUsedIterator, 1) & " Previous Used"
            intialReturnArr(emptyIndex, 2) = currentUpPreviousUsed
            intialReturnArr(emptyIndex, 3) = previousUpPreviousUsed
            intialReturnArr(emptyIndex, 4) = Result

        Next previousUsedIterator

        previousUpPreviousUsedSum = Application.Run("utilityFunction.sumArrColumn", sourceDataPreviousUpClause9, 27) / 2


        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(currentUpPreviousUsedSum, 2), Round(previousUpPreviousUsedSum, 2), 0.1)

                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(Round(currentUpPreviousUsedSum, 2) - Round(previousUpPreviousUsedSum, 2))
                End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("S" & UBound(arrUpClause9, 1)), Result
            Application.Run "EditComment", arrUpClause9Range.Range("S" & UBound(arrUpClause9, 1)), "Checked sum to this column & compare with Previous UP Used sum " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "Total Previous Used"
            intialReturnArr(emptyIndex, 2) = currentUpPreviousUsedSum
            intialReturnArr(emptyIndex, 3) = previousUpPreviousUsedSum
            intialReturnArr(emptyIndex, 4) = Result

    '    previous used raw metarial compare with previous UP end





    '    new import compare with source data import performance start

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = " "
            intialReturnArr(emptyIndex, 2) = "New Import"
            intialReturnArr(emptyIndex, 3) = ""
            intialReturnArr(emptyIndex, 4) = ""

        Dim currentUpImport, sourceDataImportPerformanceImport   As Variant

        Dim importIterator As Integer


        For importIterator = 1 To UBound(arrUpClause9, 1)

        currentUpImport = arrUpClause9(importIterator, 15)


        sourceDataImportPerformanceImport = sourceDataImportPerformanceTotalSummary(importIterator + 1, 5)

        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(currentUpImport, 2), Round(sourceDataImportPerformanceImport, 2), 0.1)

                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch = " & Round(Round(currentUpImport, 2) - Round(sourceDataImportPerformanceImport, 2), 2)
                End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause9Range.Range("O" & importIterator), Result
            Application.Run "EditComment", arrUpClause9Range.Range("O" & importIterator), "Checked with import performance statement " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = arrUpClause9(importIterator, 1) & " Import"
            intialReturnArr(emptyIndex, 2) = currentUpImport
            intialReturnArr(emptyIndex, 3) = sourceDataImportPerformanceImport
            intialReturnArr(emptyIndex, 4) = Result

        Next importIterator




    '    new import compare with source data import performance end




        Dim intialReturnArrCropIndex As Integer
        intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"


        upClause9CompareWithSource = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)


    End Function


    
    
    




Private Function upClause11CompareWithSource(arrUpClause11Range As Variant, upClause6Buyerinformation As Variant, upClause7Lcinformation As Variant, upClause7Result As Variant, sourceDataUpIssuingStatus As Variant) As Variant
    '    this function give compare result of UP clause 11 with source data
    
        Dim arrUpClause11 As Variant
        arrUpClause11 = arrUpClause11Range.value
    
        Dim regex As New RegExp
        regex.Global = True
        regex.MultiLine = True
    
    
        Dim emptyIndex As Variant
        Dim Result As Variant
    
        Dim intialReturnArr() As Variant
        ReDim intialReturnArr(1 To 100, 1 To 4)
        intialReturnArr(1, 1) = "Topic"
        intialReturnArr(1, 2) = "UP Data"
        intialReturnArr(1, 3) = "Source Data"
        intialReturnArr(1, 4) = "Result"
    
    
        intialReturnArr(2, 1) = " "
        intialReturnArr(2, 2) = "Buyer Name & UP/IP/EXP Information(UP Clause 11)"
        intialReturnArr(2, 3) = ""
        intialReturnArr(2, 4) = ""
        
        Dim buyerNameFromClause6, buyerNameFromClause11 As Variant
        
        If IsArray(upClause6Buyerinformation) Then
    
    
            Dim buyerIterator As Integer
    
            For buyerIterator = 1 To UBound(upClause6Buyerinformation, 1)
    
                buyerNameFromClause6 = upClause6Buyerinformation(buyerIterator, 1)
                buyerNameFromClause11 = arrUpClause11(buyerIterator, 3)
                
                regex.pattern = "^\d\)"
                Result = Trim(regex.Replace(Trim(buyerNameFromClause6), "")) = Trim(buyerNameFromClause11)
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch"
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause11Range.Range("C" & buyerIterator), Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Buyer"
                intialReturnArr(emptyIndex, 2) = buyerIterator & ") " & buyerNameFromClause11
                intialReturnArr(emptyIndex, 3) = buyerNameFromClause6
                intialReturnArr(emptyIndex, 4) = Result
    
    
    
            Next buyerIterator
        
    
        Else
    
    
    
                buyerNameFromClause6 = upClause6Buyerinformation
                buyerNameFromClause11 = arrUpClause11(1, 3)
                
                regex.pattern = "^\d\)"
                Result = Trim(regex.Replace(Trim(buyerNameFromClause6), "")) = Trim(buyerNameFromClause11)
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch"
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause11Range.Range("C" & 1), Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Buyer"
                intialReturnArr(emptyIndex, 2) = buyerNameFromClause11
                intialReturnArr(emptyIndex, 3) = buyerNameFromClause6
                intialReturnArr(emptyIndex, 4) = Result
    
    
        End If
        
        
            Dim asUpAllLcIndexInSourceData As Variant
            asUpAllLcIndexInSourceData = Split(upClause7Result(UBound(upClause7Result, 1), 2), " ")
                    
            Dim udExpIpIterator As Integer
            
            For udExpIpIterator = 1 To UBound(arrUpClause11, 1) - 1
                
                
                        Dim udExpIpReturnArr As Variant
    
                        udExpIpReturnArr = Application.Run("utilityFunction.mLcUdExpIpCompareWithSource", Replace(arrUpClause11(udExpIpIterator, 16), " ", ""), Replace(sourceDataUpIssuingStatus(asUpAllLcIndexInSourceData(udExpIpIterator - 1), 17), " ", ""), sourceDataUpIssuingStatus(asUpAllLcIndexInSourceData(udExpIpIterator - 1), 18), "UD/IP/EXP")
    
                        intialReturnArr = Application.Run("utilityFunction.mergeArry", intialReturnArr, udExpIpReturnArr, 1)
                        
                        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause11Range.Range("p" & udExpIpIterator), Application.Run("utilityFunction.isAllResultOk", udExpIpReturnArr)
                    
            Next udExpIpIterator
    
        
        
        Dim clause11QtySum As Variant
        
        clause11QtySum = Application.Run("utilityFunction.sumArrColumn", arrUpClause11, 25) - arrUpClause11(UBound(arrUpClause11, 1), 25)
        
        Result = Round(clause11QtySum, 2) = Round(arrUpClause11(UBound(arrUpClause11, 1), 25), 2)
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause11QtySum, 2) - Round(arrUpClause11(UBound(arrUpClause11, 1), 25), 2), 2)
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause11Range.Range("Y1:Y" & UBound(arrUpClause11, 1) - 1), Result
                
                Application.Run "EditComment", arrUpClause11Range.Range("Y" & UBound(arrUpClause11, 1)), "Checked sum to this column " & Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Qty. Sum"
                intialReturnArr(emptyIndex, 2) = clause11QtySum
                intialReturnArr(emptyIndex, 3) = arrUpClause11(UBound(arrUpClause11, 1), 25)
                intialReturnArr(emptyIndex, 4) = Result
                
                
                
        Dim clause11TotalQty, clause7TotalQty As Variant
        
        clause7TotalQty = upClause7Lcinformation(UBound(upClause7Lcinformation, 1), 17)
        
        clause11TotalQty = arrUpClause11(UBound(arrUpClause11, 1), 25)
        
        Result = Round(clause11TotalQty, 2) = Round(clause7TotalQty, 2)
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause11TotalQty, 2) - Round(clause7TotalQty, 2), 2)
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause11Range.Range("Y" & UBound(arrUpClause11, 1)), Result
                
                Application.Run "EditComment", arrUpClause11Range.Range("Y" & UBound(arrUpClause11, 1)), "Checked with UP clause 7 " & Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Total Qty."
                intialReturnArr(emptyIndex, 2) = clause11TotalQty & " (Clause 11)"
                intialReturnArr(emptyIndex, 3) = clause7TotalQty & " (Clause 7)"
                intialReturnArr(emptyIndex, 4) = Result
        
        
    
        
        Dim intialReturnArrCropIndex As Integer
        intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"
    
    
        upClause11CompareWithSource = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)
    
    
    End Function
    
 



Private Function upClause12aCompareWithSource(arrUpClause12aRange As Variant, upClause6Buyerinformation As Variant, upClause7Lcinformation As Variant, upYarnConsumptionInformation As Variant) As Variant
    '    this function give compare result of UP clause 12a with source data
    
        Dim arrUpClause12a As Variant
        arrUpClause12a = arrUpClause12aRange.value
    
        Dim regex As New RegExp
        regex.Global = True
        regex.MultiLine = True
    
    
        Dim emptyIndex As Variant
        Dim Result As Variant
    
        Dim intialReturnArr(1 To 50, 1 To 4) As Variant
        intialReturnArr(1, 1) = "Topic"
        intialReturnArr(1, 2) = "UP Data"
        intialReturnArr(1, 3) = "Source Data"
        intialReturnArr(1, 4) = "Result"
    
    
        intialReturnArr(2, 1) = " "
        intialReturnArr(2, 2) = "Yarn Consumption Information(UP Clause 12a)"
        intialReturnArr(2, 3) = ""
        intialReturnArr(2, 4) = ""
        
        ' buyer compare start
    
        Dim buyerArrFromClause12a, arrUpClause12aCroped As Variant
        arrUpClause12aCroped = Application.Run("utilityFunction.cropedArry", arrUpClause12a, 1, UBound(arrUpClause12a, 1) - 1)
        buyerArrFromClause12a = Application.Run("utilityFunction.towDimensionalArrayFilter", arrUpClause12aCroped, "[a-zA-Z]", 2)
    
        Dim buyerNameFromClause6, buyerNameFromClause12a As Variant
        
        If IsArray(upClause6Buyerinformation) Then
    
    
            Dim buyerIterator As Integer
    
            For buyerIterator = 1 To UBound(upClause6Buyerinformation, 1)
    
                buyerNameFromClause6 = upClause6Buyerinformation(buyerIterator, 1)
                buyerNameFromClause12a = buyerArrFromClause12a(buyerIterator, 2)
                
                regex.pattern = "^\d\)"
                Result = Trim(regex.Replace(Trim(buyerNameFromClause6), "")) = Trim(buyerNameFromClause12a)
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch"
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12aRange.Range("B" & buyerArrFromClause12a(buyerIterator, UBound(buyerArrFromClause12a, 2))), Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Buyer"
                intialReturnArr(emptyIndex, 2) = buyerIterator & ") " & buyerNameFromClause12a
                intialReturnArr(emptyIndex, 3) = buyerNameFromClause6
                intialReturnArr(emptyIndex, 4) = Result
    
    
    
            Next buyerIterator
        
    
        Else
    
    
    
                buyerNameFromClause6 = upClause6Buyerinformation
                buyerNameFromClause12a = buyerArrFromClause12a(1, 2)
                
                regex.pattern = "^\d\)"
                Result = Trim(regex.Replace(Trim(buyerNameFromClause6), "")) = Trim(buyerNameFromClause12a)
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch"
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12aRange.Range("B" & 1), Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Buyer"
                intialReturnArr(emptyIndex, 2) = buyerNameFromClause12a
                intialReturnArr(emptyIndex, 3) = buyerNameFromClause6
                intialReturnArr(emptyIndex, 4) = Result
    
    
        End If
    
    ' buyer compare end
        
        
        
        Dim clause12aQtySum As Variant
        
        clause12aQtySum = Application.Run("utilityFunction.sumArrColumn", arrUpClause12a, 18) - arrUpClause12a(UBound(arrUpClause12a, 1), 18)
        
        Result = Round(clause12aQtySum, 2) = Round(arrUpClause12a(UBound(arrUpClause12a, 1), 18), 2)
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12aQtySum, 2) - Round(arrUpClause12a(UBound(arrUpClause12a, 1), 18), 2), 2)
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12aRange.Range("R1:R" & UBound(arrUpClause12a, 1) - 1), Result
                
                Application.Run "EditComment", arrUpClause12aRange.Range("R" & UBound(arrUpClause12a, 1)), "Checked sum to this column " & Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Qty. Sum"
                intialReturnArr(emptyIndex, 2) = clause12aQtySum
                intialReturnArr(emptyIndex, 3) = arrUpClause12a(UBound(arrUpClause12a, 1), 18)
                intialReturnArr(emptyIndex, 4) = Result
    
    
    
        Dim clause12aUsedYarnQtySum As Variant
        
        clause12aUsedYarnQtySum = Application.Run("utilityFunction.sumArrColumn", arrUpClause12a, 25) - arrUpClause12a(UBound(arrUpClause12a, 1), 25)
        
        Result = Round(clause12aUsedYarnQtySum, 2) = Round(arrUpClause12a(UBound(arrUpClause12a, 1), 25), 2)
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12aUsedYarnQtySum, 2) - Round(arrUpClause12a(UBound(arrUpClause12a, 1), 25), 2), 2)
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12aRange.Range("Y1:Y" & UBound(arrUpClause12a, 1) - 1), Result
                
                Application.Run "EditComment", arrUpClause12aRange.Range("Y" & UBound(arrUpClause12a, 1)), "Checked sum to this column " & Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Used Yarn Qty. Sum"
                intialReturnArr(emptyIndex, 2) = clause12aUsedYarnQtySum
                intialReturnArr(emptyIndex, 3) = arrUpClause12a(UBound(arrUpClause12a, 1), 25)
                intialReturnArr(emptyIndex, 4) = Result
    
    
        Dim clause12aTotalQty, clause7TotalQty As Variant
        
        clause7TotalQty = upClause7Lcinformation(UBound(upClause7Lcinformation, 1), 17)
        
        clause12aTotalQty = arrUpClause12a(UBound(arrUpClause12a, 1), 18)
        
        Result = Round(clause12aTotalQty, 2) = Round(clause7TotalQty, 2)
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12aTotalQty, 2) - Round(clause7TotalQty, 2), 2)
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12aRange.Range("R" & UBound(arrUpClause12a, 1)), Result
                
                Application.Run "EditComment", arrUpClause12aRange.Range("R" & UBound(arrUpClause12a, 1)), "Checked with UP clause 7 " & Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Total Qty."
                intialReturnArr(emptyIndex, 2) = clause12aTotalQty & " (Clause 12a)"
                intialReturnArr(emptyIndex, 3) = clause7TotalQty & " (Clause 7)"
                intialReturnArr(emptyIndex, 4) = Result
    
    
        Dim clause12aTotalUsedYarnQty, totalConsumptionYarnQty As Variant
        
        totalConsumptionYarnQty = upYarnConsumptionInformation(1, 8)
        
        clause12aTotalUsedYarnQty = arrUpClause12a(UBound(arrUpClause12a, 1), 25)
        
        Result = Round(clause12aTotalUsedYarnQty, 2) = Round(totalConsumptionYarnQty, 2)
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12aTotalUsedYarnQty, 2) - Round(totalConsumptionYarnQty, 2), 2)
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12aRange.Range("Y" & UBound(arrUpClause12a, 1)), Result
                
                Application.Run "EditComment", arrUpClause12aRange.Range("Y" & UBound(arrUpClause12a, 1)), "Checked with consumption sheet " & Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Total Used Yarn Qty."
                intialReturnArr(emptyIndex, 2) = clause12aTotalUsedYarnQty & " (Clause 12a)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionYarnQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result
        
        
    
        
        Dim intialReturnArrCropIndex As Integer
        intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"
    
    
        upClause12aCompareWithSource = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)
    
    
    End Function
    
    
Private Function upClause12bCompareWithSource(arrUpClause12bRange As Variant, upClause6Buyerinformation As Variant, upYarnConsumptionInformation As Variant, upClause7Buyerinformation As Variant) As Variant
    '    this function give compare result of UP clause 12b with source data

        Dim arrUpClause12b As Variant
        arrUpClause12b = arrUpClause12bRange.value

        Dim regex As New RegExp
        regex.Global = True
        regex.MultiLine = True


        Dim emptyIndex As Variant
        Dim Result As Variant

        Dim intialReturnArr(1 To 50, 1 To 4) As Variant
        intialReturnArr(1, 1) = "Topic"
        intialReturnArr(1, 2) = "UP Data"
        intialReturnArr(1, 3) = "Source Data"
        intialReturnArr(1, 4) = "Result"


        intialReturnArr(2, 1) = " "
        intialReturnArr(2, 2) = "Chemical & Dyes Consumption Information(UP Clause 12b)"
        intialReturnArr(2, 3) = ""
        intialReturnArr(2, 4) = ""

        ' buyer compare start

        Dim buyerArrFromClause12b As Variant

        buyerArrFromClause12b = Application.Run("utilityFunction.towDimensionalArrayFilter", arrUpClause12b, "[a-zA-Z]", 1)
        buyerArrFromClause12b = Application.Run("utilityFunction.cropedArry", buyerArrFromClause12b, 2, UBound(buyerArrFromClause12b, 1)) ' exclude row 1

        Dim buyerNameFromClause6, buyerNameFromClause12b As Variant

        If IsArray(upClause6Buyerinformation) Then


            Dim buyerIterator As Integer

            For buyerIterator = 1 To UBound(upClause6Buyerinformation, 1)

                buyerNameFromClause6 = upClause6Buyerinformation(buyerIterator, 1)
                buyerNameFromClause12b = buyerArrFromClause12b(buyerIterator, 1)

                regex.pattern = "^\d\)"
                Result = Trim(regex.Replace(Trim(buyerNameFromClause6), "")) = Trim(buyerNameFromClause12b)

                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch"
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("A" & buyerArrFromClause12b(buyerIterator, UBound(buyerArrFromClause12b, 2))), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Buyer"
                intialReturnArr(emptyIndex, 2) = buyerIterator & ") " & buyerNameFromClause12b
                intialReturnArr(emptyIndex, 3) = buyerNameFromClause6
                intialReturnArr(emptyIndex, 4) = Result



            Next buyerIterator


        Else



                buyerNameFromClause6 = upClause6Buyerinformation
                buyerNameFromClause12b = buyerArrFromClause12b(1, 1)

                regex.pattern = "^\d\)"
                Result = Trim(regex.Replace(Trim(buyerNameFromClause6), "")) = Trim(buyerNameFromClause12b)

                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch"
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("A" & 2), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Buyer"
                intialReturnArr(emptyIndex, 2) = buyerNameFromClause12b
                intialReturnArr(emptyIndex, 3) = buyerNameFromClause6
                intialReturnArr(emptyIndex, 4) = Result


        End If

    ' buyer compare end



    ' total yarn Qty compare start
        Dim clause12bTotalYarnQty, totalConsumptionYarnQty As Variant

        totalConsumptionYarnQty = upYarnConsumptionInformation(1, 8)

        clause12bTotalYarnQty = arrUpClause12b(1, 6)

        Result = Round(clause12bTotalYarnQty, 2) = Round(totalConsumptionYarnQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYarnQty, 2) - Round(totalConsumptionYarnQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 1), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Total Yarn Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYarnQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionYarnQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total yarn Qty compare end




    ' total black Qty compare start
        Dim clause12bTotalYBlackQty, totalConsumptionBlackQty As Variant

        totalConsumptionBlackQty = upYarnConsumptionInformation(3, 8)

        clause12bTotalYBlackQty = arrUpClause12b(12, 6)

        Result = Round(clause12bTotalYBlackQty, 2) = Round(totalConsumptionBlackQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYBlackQty, 2) - Round(totalConsumptionBlackQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 12), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Black Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYBlackQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionBlackQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total black Qty compare end


    ' total Mercerization(Sulphur) Qty compare start
        Dim clause12bTotalYMercerizationSulphurQty, totalConsumptionMercerizationSulphurkQty As Variant

        totalConsumptionMercerizationSulphurkQty = upYarnConsumptionInformation(4, 8)

        clause12bTotalYMercerizationSulphurQty = arrUpClause12b(18, 6)

        Result = Round(clause12bTotalYMercerizationSulphurQty, 2) = Round(totalConsumptionMercerizationSulphurkQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYMercerizationSulphurQty, 2) - Round(totalConsumptionMercerizationSulphurkQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 18), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Mercerization Black Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYMercerizationSulphurQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionMercerizationSulphurkQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total Mercerization(Sulphur) Qty compare end



    ' total indigo Qty compare start
        Dim clause12bTotalYIndigoQty, totalConsumptionIndigoQty As Variant

        totalConsumptionIndigoQty = upYarnConsumptionInformation(5, 8)

        clause12bTotalYIndigoQty = arrUpClause12b(32, 6)

        Result = Round(clause12bTotalYIndigoQty, 2) = Round(totalConsumptionIndigoQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYIndigoQty, 2) - Round(totalConsumptionIndigoQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 32), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Indigo Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYIndigoQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionIndigoQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total indigo Qty compare end


    ' total Mercerization(Indigo) Qty compare start
        Dim clause12bTotalYMercerizationIndigoQty, totalConsumptionMercerizationIndigoQty As Variant

        totalConsumptionMercerizationIndigoQty = upYarnConsumptionInformation(6, 8)

        clause12bTotalYMercerizationIndigoQty = arrUpClause12b(38, 6)

        Result = Round(clause12bTotalYMercerizationIndigoQty, 2) = Round(totalConsumptionMercerizationIndigoQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYMercerizationIndigoQty, 2) - Round(totalConsumptionMercerizationIndigoQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 38), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Mercerization Indigo Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYMercerizationIndigoQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionMercerizationIndigoQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total Mercerization(Indigo) Qty compare end




    ' total dyed Qty compare start
        Dim clause12bTotalYDyedQty, totalConsumptionDyedQty As Variant

        totalConsumptionDyedQty = upYarnConsumptionInformation(7, 8)

        clause12bTotalYDyedQty = arrUpClause12b(53, 6)

        Result = Round(clause12bTotalYDyedQty, 2) = Round(totalConsumptionDyedQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYDyedQty, 2) - Round(totalConsumptionDyedQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 53), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Dyed Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYDyedQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionDyedQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total dyed Qty compare end



    ' total Mercerization(Dyed) Qty compare start
        Dim clause12bTotalYMercerizationDyedQty, totalConsumptionMercerizationDyedQty As Variant

        totalConsumptionMercerizationDyedQty = upYarnConsumptionInformation(8, 8)

        clause12bTotalYMercerizationDyedQty = arrUpClause12b(61, 6)

        Result = Round(clause12bTotalYMercerizationDyedQty, 2) = Round(totalConsumptionMercerizationDyedQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYMercerizationDyedQty, 2) - Round(totalConsumptionMercerizationDyedQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 61), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Mercerization Dyed Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYMercerizationDyedQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionMercerizationDyedQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total Mercerization(Dyed) Qty compare end


    ' total Over Dying Qty compare start
        Dim clause12bTotalYOverDyingQty, totalConsumptionOverDyingQty As Variant

        totalConsumptionOverDyingQty = upYarnConsumptionInformation(9, 8)

        clause12bTotalYOverDyingQty = arrUpClause12b(72, 6)

        Result = Round(clause12bTotalYOverDyingQty, 2) = Round(totalConsumptionOverDyingQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYOverDyingQty, 2) - Round(totalConsumptionOverDyingQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 72), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Over Dying Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYOverDyingQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionOverDyingQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total Over Dying Qty compare end



        ' total Mercerization(Over Dying) Qty compare start
        Dim clause12bTotalYMercerizationOverDyingQty, totalConsumptionMercerizationOverDyingQty As Variant

        totalConsumptionMercerizationOverDyingQty = upYarnConsumptionInformation(10, 8)

        clause12bTotalYMercerizationOverDyingQty = arrUpClause12b(75, 6)

        Result = Round(clause12bTotalYMercerizationOverDyingQty, 2) = Round(totalConsumptionMercerizationOverDyingQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYMercerizationOverDyingQty, 2) - Round(totalConsumptionMercerizationOverDyingQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 75), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Mercerization Over Dying Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYMercerizationOverDyingQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionMercerizationOverDyingQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total Mercerization(Over Dying) Qty compare end


    ' total Coating Qty compare start
        Dim clause12bTotalYCoatingQty, totalConsumptionCoatingQty As Variant

        totalConsumptionCoatingQty = upYarnConsumptionInformation(11, 8)

        clause12bTotalYCoatingQty = arrUpClause12b(82, 6)

        Result = Round(clause12bTotalYCoatingQty, 2) = Round(totalConsumptionCoatingQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYCoatingQty, 2) - Round(totalConsumptionCoatingQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 82), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Coating Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYCoatingQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionCoatingQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total Coating Qty compare end



        ' total PFD Qty compare start
        Dim clause12bTotalYPFDQty, totalConsumptionPFDQty As Variant

        totalConsumptionPFDQty = upYarnConsumptionInformation(12, 8)

        clause12bTotalYPFDQty = arrUpClause12b(92, 6)

        Result = Round(clause12bTotalYPFDQty, 2) = Round(totalConsumptionPFDQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYPFDQty, 2) - Round(totalConsumptionPFDQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 92), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "PFD Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYPFDQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionPFDQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total PFD Qty compare end


        ' total ECRU Qty compare start
        Dim clause12bTotalYECRUQty, totalConsumptionECRUQty As Variant

        totalConsumptionECRUQty = upYarnConsumptionInformation(13, 8)

        clause12bTotalYECRUQty = arrUpClause12b(102, 6)

        Result = Round(clause12bTotalYECRUQty, 2) = Round(totalConsumptionECRUQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bTotalYECRUQty, 2) - Round(totalConsumptionECRUQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 102), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "ECRU Qty."
                intialReturnArr(emptyIndex, 2) = clause12bTotalYECRUQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionECRUQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' total ECRU Qty compare end


        ' ETP & WTP Qty compare start
        Dim clause12bETPandWTPQty, totalConsumptionETPandWTPQty As Variant

        totalConsumptionETPandWTPQty = upYarnConsumptionInformation(1, 8)

        clause12bETPandWTPQty = arrUpClause12b(107, 6)

        Result = Round(clause12bETPandWTPQty, 2) = Round(totalConsumptionETPandWTPQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bETPandWTPQty, 2) - Round(totalConsumptionETPandWTPQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 107), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "ETP & WTP Qty."
                intialReturnArr(emptyIndex, 2) = clause12bETPandWTPQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalConsumptionETPandWTPQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' ETP & WTP Qty compare end


        ' Packing Qty compare start
        Dim clause12bPackingQty, totalFabricPackingQty As Variant

        totalFabricPackingQty = upClause7Buyerinformation(UBound(upClause7Buyerinformation, 1), 17)

        clause12bPackingQty = arrUpClause12b(111, 6)

        Result = Round(clause12bPackingQty, 2) = Round(totalFabricPackingQty, 2)


                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(clause12bPackingQty, 2) - Round(totalFabricPackingQty, 2), 2)
                    End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause12bRange.Range("F" & 111), Result

                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Packing Qty."
                intialReturnArr(emptyIndex, 2) = clause12bPackingQty & " (Clause 12b)"
                intialReturnArr(emptyIndex, 3) = totalFabricPackingQty & " (Yarn Consumption Sheet)"
                intialReturnArr(emptyIndex, 4) = Result

    ' Packing Qty compare end


        Dim intialReturnArrCropIndex As Integer
        intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"


        upClause12bCompareWithSource = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)


    End Function





Private Function upClause13CompareWithSource(arrUpClause13Range As Variant, resultClause8 As Variant) As Variant
    '    this function give compare result of UP clause 13 with source data
    
        Dim arrUpClause13 As Variant
        arrUpClause13 = arrUpClause13Range.value
    
        Dim regex As New RegExp
        regex.Global = True
        regex.MultiLine = True
    
    
        Dim emptyIndex As Variant
        Dim Result As Variant
    
        Dim intialReturnArr(1 To 50, 1 To 4) As Variant
        intialReturnArr(1, 1) = "Topic"
        intialReturnArr(1, 2) = "UP Data"
        intialReturnArr(1, 3) = "Source Data"
        intialReturnArr(1, 4) = "Result"
    
    
        intialReturnArr(2, 1) = " "
        intialReturnArr(2, 2) = "Used Raw Materials Information(UP Clause 13)"
        intialReturnArr(2, 3) = ""
        intialReturnArr(2, 4) = ""
        
    
    
    
    
    
    
        Dim resultClause8Croped As Variant
        Dim resultClause8CropedOdd, resultClause8CropedEven As Variant
    
    
        resultClause8Croped = Application.Run("utilityFunction.cropedArry", resultClause8, 8, 19)
    
        resultClause8CropedOdd = Application.Run("utilityFunction.evenOrOddIndexArrayFilter", resultClause8Croped, "odd", False)
        resultClause8CropedEven = Application.Run("utilityFunction.evenOrOddIndexArrayFilter", resultClause8Croped, "even", False)
        
    
    
        '    Qty compare start
    
        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
        intialReturnArr(emptyIndex, 1) = " "
        intialReturnArr(emptyIndex, 2) = "Qty"
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""
    
        Dim qtyIterator As Integer
    
        Dim qtyFromUpClause13, qtyFromResultClause8 As Variant
        
        Dim sumQty As Variant
        
        sumQty = 0
        
        For qtyIterator = 1 To 6
    
            qtyFromUpClause13 = arrUpClause13(qtyIterator, 14)
            qtyFromResultClause8 = resultClause8CropedOdd(qtyIterator, 2)
            
            sumQty = sumQty + qtyFromUpClause13
    
            Result = Round(qtyFromUpClause13, 2) = Round(qtyFromResultClause8, 2)
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(qtyFromUpClause13, 2) - Round(qtyFromResultClause8, 2), 2)
                    End If
    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause13Range.Range("N" & qtyIterator), Result
                Application.Run "EditComment", arrUpClause13Range.Range("N" & qtyIterator), "Checked with UP clause 8 " & Result
    
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = arrUpClause13(qtyIterator, 4) & " Qty"
                intialReturnArr(emptyIndex, 2) = qtyFromUpClause13 & " (Clause 13)"
                intialReturnArr(emptyIndex, 3) = qtyFromResultClause8 & " (Clause 8)"
                intialReturnArr(emptyIndex, 4) = Result
    
    
            
        Next qtyIterator
        
        
            Result = Round(sumQty, 2) = Round(arrUpClause13(7, 14), 2)
            
            Dim qtySumResult As Boolean
            qtySumResult = Result
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(sumQty, 2) - Round(arrUpClause13(7, 14), 2), 2)
                    End If
    
                Application.Run "EditComment", arrUpClause13Range.Range("N" & 7), "Checked sum to this column " & Result
                
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Total Qty Sum"
                intialReturnArr(emptyIndex, 2) = sumQty
                intialReturnArr(emptyIndex, 3) = arrUpClause13(7, 14)
                intialReturnArr(emptyIndex, 4) = Result
                
                
                
                
                
                
            Result = Round(arrUpClause13(7, 14), 2) = Round(resultClause8(4, 2), 2)
            
            Dim totalQtyResult As Boolean
            totalQtyResult = Result
                
    
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(Round(arrUpClause13(7, 14), 2) - Round(resultClause8(4, 2), 2), 2)
                    End If
    
                Application.Run "EditComment", arrUpClause13Range.Range("N" & 7), "Checked with UP clause 8 " & Result
                
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = "Total Qty"
                intialReturnArr(emptyIndex, 2) = arrUpClause13(7, 14) & " (Clause 13)"
                intialReturnArr(emptyIndex, 3) = resultClause8(4, 2) & " (Clause 8)"
                intialReturnArr(emptyIndex, 4) = Result
                
                
               Result = qtySumResult = True And totalQtyResult = True
               
               If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch"
                End If
                    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause13Range.Range("N" & 7), Result
                
        
        
        '    Qty compare end
    
    
    
    
          '    value compare start
        
            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
            intialReturnArr(emptyIndex, 1) = " "
            intialReturnArr(emptyIndex, 2) = "Value"
            intialReturnArr(emptyIndex, 3) = ""
            intialReturnArr(emptyIndex, 4) = ""
        
        
            Dim valueIterator As Integer
        
            Dim valueFromUpClause13, valueFromResultClause8 As Variant
            
            Dim sumValue As Variant
            
            sumValue = 0
            
            For valueIterator = 1 To 6
        
                valueFromUpClause13 = arrUpClause13(valueIterator, 17)
                valueFromResultClause8 = resultClause8CropedEven(valueIterator, 2)
                
                sumValue = sumValue + valueFromUpClause13
        
                Result = Round(valueFromUpClause13, 2) = Round(valueFromResultClause8, 2)
                    
        
                        If Result Then
                            Result = "OK"
                        Else
                            Result = "Mismatch = " & Round(Round(valueFromUpClause13, 2) - Round(valueFromResultClause8, 2), 2)
                        End If
        
                    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause13Range.Range("Q" & valueIterator), Result
                    Application.Run "EditComment", arrUpClause13Range.Range("Q" & valueIterator), "Checked with UP clause 8 " & Result
        
                    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
                    intialReturnArr(emptyIndex, 1) = arrUpClause13(valueIterator, 4) & " Value"
                    intialReturnArr(emptyIndex, 2) = valueFromUpClause13 & " (Clause 13)"
                    intialReturnArr(emptyIndex, 3) = valueFromResultClause8 & " (Clause 8)"
                    intialReturnArr(emptyIndex, 4) = Result
        
        
                
            Next valueIterator
            
            
                Result = Round(sumValue, 2) = Round(arrUpClause13(7, 17), 2)
                
                Dim valueSumResult As Boolean
                valueSumResult = Result
        
                        If Result Then
                            Result = "OK"
                        Else
                            Result = "Mismatch = " & Round(Round(sumValue, 2) - Round(arrUpClause13(7, 17), 2), 2)
                        End If
        
                    Application.Run "EditComment", arrUpClause13Range.Range("Q" & 7), "Checked sum to this column " & Result
                    
                    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
                    intialReturnArr(emptyIndex, 1) = "Total Value Sum"
                    intialReturnArr(emptyIndex, 2) = sumValue
                    intialReturnArr(emptyIndex, 3) = arrUpClause13(7, 17)
                    intialReturnArr(emptyIndex, 4) = Result
                    
                    
                    
                Result = Round(arrUpClause13(7, 17), 2) = Round(resultClause8(5, 2), 2)
                
                Dim totalValueResult As Boolean
                totalValueResult = Result
                    
        
                        If Result Then
                            Result = "OK"
                        Else
                            Result = "Mismatch = " & Round(Round(arrUpClause13(7, 17), 2) - Round(resultClause8(5, 2), 2), 2)
                        End If
        
                    Application.Run "EditComment", arrUpClause13Range.Range("Q" & 7), "Checked with UP clause 8 " & Result
                    
                    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
                    intialReturnArr(emptyIndex, 1) = "Total Value"
                    intialReturnArr(emptyIndex, 2) = arrUpClause13(7, 17) & " (Clause 13)"
                    intialReturnArr(emptyIndex, 3) = resultClause8(5, 2) & " (Clause 8)"
                    intialReturnArr(emptyIndex, 4) = Result
                    
                    
                   Result = valueSumResult = True And totalValueResult = True
                   
                   If Result Then
                            Result = "OK"
                        Else
                            Result = "Mismatch"
                    End If
                        
                    Application.Run "utilityFunction.errorMarkingForValue", arrUpClause13Range.Range("Q" & 7), Result
                    
            
            
            '    value compare end
    
    
    
    
        
        Dim intialReturnArrCropIndex As Integer
        intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"
    
    
        upClause13CompareWithSource = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)
    
    
    End Function
    

