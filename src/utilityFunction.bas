Attribute VB_Name = "utilityFunction"
Option Explicit

Private Function towDimensionalArrayFilter(arr() As Variant, patternStr As String, filterIndex As Integer) As Variant
'this function give filtered array from source array
    
    
    Dim regex As New RegExp
    regex.Global = True
    regex.pattern = patternStr
    regex.MultiLine = True
    
    Dim innerArrLength As Integer
    innerArrLength = UBound(arr, 2)
     
    Dim returnArrLength As Integer ' Dynamic Multidimensional Array, reDim work only last dimension
    returnArrLength = 0
    Dim i As Integer
    For i = 1 To UBound(arr, 1)
        If regex.test(arr(i, filterIndex)) Then
            returnArrLength = returnArrLength + 1
        End If
    Next i
    
    If returnArrLength = 0 Then
        towDimensionalArrayFilter = Null
        Exit Function
    End If
    
    Dim returnArr() As Variant
    ReDim returnArr(1 To returnArrLength, 1 To innerArrLength + 1) ' add 1 to inner arr length for store original index no.
    
    Dim j, k, counter As Integer
    counter = 0
    For j = 1 To UBound(arr, 1)
        If regex.test(arr(j, filterIndex)) Then
            counter = counter + 1
           For k = 1 To UBound(arr, 2)
            returnArr(counter, k) = arr(j, k)
           Next k
           returnArr(counter, innerArrLength + 1) = j ' push original index to last column for next time use if any
        End If
    Next j
    
    towDimensionalArrayFilter = returnArr

End Function





Private Function towDimensionalArrayFilterNegative(arr() As Variant, patternStr As String, filterIndex As Integer) As Variant
'this function give filtered array from source array
    
    
    Dim regex As New RegExp
    regex.Global = True
    regex.pattern = patternStr
    regex.MultiLine = True
    
    Dim innerArrLength As Integer
    innerArrLength = UBound(arr, 2)
     
    Dim returnArrLength As Integer ' Dynamic Multidimensional Array, reDim work only last dimension
    returnArrLength = 0
    Dim i As Integer
    For i = 1 To UBound(arr, 1)
        If Not regex.test(arr(i, filterIndex)) Then
            returnArrLength = returnArrLength + 1
        End If
    Next i
    
    If returnArrLength = 0 Then
        towDimensionalArrayFilterNegative = Null
        Exit Function
    End If
    
    Dim returnArr() As Variant
    ReDim returnArr(1 To returnArrLength, 1 To innerArrLength + 1) ' add 1 to inner arr length for store original index no.
    
    Dim j, k, counter As Integer
    counter = 0
    For j = 1 To UBound(arr, 1)
        If Not regex.test(arr(j, filterIndex)) Then
            counter = counter + 1
           For k = 1 To UBound(arr, 2)
            returnArr(counter, k) = arr(j, k)
           Next k
           returnArr(counter, innerArrLength + 1) = j ' push original index to last column for next time use if any
        End If
    Next j
    
    towDimensionalArrayFilterNegative = returnArr

End Function








Private Function towDimensionalArrayFilterWithNextIndex(arr() As Variant, patternStr As String, filterIndex As Integer) As Variant
'this function give filtered array(note: when meet filter criteria then also filter current & next row, it's help to filter Bill of Entry/ Mushak & their Qty.) from source array
    
    
    If UBound(arr, 1) Mod 2 <> 0 Then ' validation
    towDimensionalArrayFilterWithNextIndex = Null
        Exit Function
    End If
    
    
    Dim regex As New RegExp
    regex.Global = True
    regex.pattern = patternStr
    regex.MultiLine = True
    
    Dim innerArrLength As Integer
    innerArrLength = UBound(arr, 2)
     
    Dim returnArrLength As Integer ' Dynamic Multidimensional Array, reDim work only last dimension
    returnArrLength = 0
    Dim i As Integer
    For i = 1 To UBound(arr, 1)
        If regex.test(arr(i, filterIndex)) Then
            returnArrLength = returnArrLength + 1
        End If
    Next i
    
    If returnArrLength = 0 Then
        towDimensionalArrayFilterWithNextIndex = Null
        Exit Function
    End If
    
    Dim returnArr() As Variant
    ReDim returnArr(1 To returnArrLength * 2, 1 To innerArrLength + 1) ' add 1 to inner arr length for store original index no.
    
    Dim j, k, l, counter As Integer
    counter = 0
    For j = 1 To UBound(arr, 1)
        If regex.test(arr(j, filterIndex)) Then
            counter = counter + 1
           For k = 1 To UBound(arr, 2)
            returnArr(counter, k) = arr(j, k)
           Next k
           returnArr(counter, innerArrLength + 1) = j ' push original index to last column for next time use if any
           
           
           
            counter = counter + 1
           For l = 1 To UBound(arr, 2)
            returnArr(counter, l) = arr(j + 1, l)
           Next l
           returnArr(counter, innerArrLength + 1) = j + 1 ' push original index to last column for next time use if any
           
        End If
    Next j
    
    towDimensionalArrayFilterWithNextIndex = returnArr

End Function



Private Function openFile(fileName As String) As Variant ' provide source file name
'this function open a specific file

    Dim path As String
    path = ActiveWorkbook.path & Application.PathSeparator ' dynamic
    Workbooks.Open fileName:=path & fileName, ReadOnly:=True

End Function

Private Function openFileFullPath(filePath As String) ' provide source file full path
    'open a specific file

    Workbooks.Open fileName:=filePath, ReadOnly:=True

End Function

Private Function closeFile(fileName As String) As Variant ' provide source file name
'this function close a specific file

    Workbooks(fileName).Close SaveChanges:=False

End Function


Private Function indexOf(arr() As Variant, patternStr As String, columnIndex As Integer, startingIndex As Integer, endingIndex As Integer) As Variant ' provide source array, patternStr, criteriaColumn, starting index & ending index example # UBound(arr, 1) #
'this function give index
    
    Dim regex As New RegExp
    regex.Global = True
    regex.pattern = patternStr
    regex.MultiLine = True
    
    Dim i As Integer
    For i = startingIndex To endingIndex
        If regex.test(arr(i, columnIndex)) Then
            indexOf = i
            Exit Function
        End If
    Next i
    
    indexOf = Null

End Function

Private Function indexOfReverseOrder(arr() As Variant, patternStr As String, columnIndex As Integer, startingIndex As Integer, endingIndex As Integer) As Variant ' provide source array, patternStr, criteriaColumn, starting index example # UBound(arr, 1) # & ending index
'this function give index number from reverse order
    
    Dim regex As New RegExp
    regex.Global = True
    regex.pattern = patternStr
    regex.MultiLine = True

    Dim i As Integer
    For i = startingIndex To endingIndex Step -1
        If regex.test(arr(i, columnIndex)) Then
            indexOfReverseOrder = i
            Exit Function
        End If
    Next i

    indexOfReverseOrder = Null

End Function


Private Function sumArrColumn(arr() As Variant, columnIndex As Integer) As Variant
'this function give sum of specific column

    Dim sum As Variant
    sum = 0
    Dim i As Integer
    For i = 1 To UBound(arr, 1)
    
      sum = sum + arr(i, columnIndex)
      
    Next i
    
    sumArrColumn = sum
    
End Function


Private Function sumQtyFromDictFormat(sourceDataAsDicUpIssuingStatus As Object) As Variant
'this function give sum of Qty. & it's deticated only for sum quantity
'if any qty. unit are in Mtr. then convert to Yds and sum

    Dim sum As Variant
    sum = 0

    Dim dictKey As Variant

    For Each dictKey In sourceDataAsDicUpIssuingStatus.keys
        If Right(sourceDataAsDicUpIssuingStatus(dictKey)("qtyNumberFormat"), 5) = """Mtr""" Then

            sum = sum + Round(sourceDataAsDicUpIssuingStatus(dictKey)("QuantityofFabricsYdsMtr") * 1.0936132983, 2)

        Else

            sum = sum + sourceDataAsDicUpIssuingStatus(dictKey)("QuantityofFabricsYdsMtr")

        End If

    Next dictKey

    sumQtyFromDictFormat = Round(sum)

End Function


Private Function sumQty(arr() As Variant, sumColumnIndex As Integer, criteriaColumnIndex As Integer) As Variant
'this function give sum of Qty. & it's deticated only for sum quantity
'if any qty. unit are in Mtr. then convert to Yds and sum


    Dim sum As Variant
    sum = 0
    Dim i As Integer
    For i = 1 To UBound(arr, 1)
        If arr(i, criteriaColumnIndex) <> "Mtr" Then
        
            sum = sum + arr(i, sumColumnIndex)
            
        Else
        
            sum = sum + Round(arr(i, sumColumnIndex) * 1.0936132983)
        
        End If
      
    Next i
    
    sumQty = sum
    
End Function


Private Function evenOrOddIndexArrayFilter(arr() As Variant, evenOrOdd As String, lengthValidation As Boolean) As Variant ' provide sourch arr, (evenOrOdd = "even" or "odd")
'this function give even or odd index filtered array from source array

    
    If lengthValidation Then
        If UBound(arr, 1) Mod 2 <> 0 Then ' validation
        evenOrOddIndexArrayFilter = Null
            Exit Function
        End If
    End If
    
    Dim modResult As Integer
    
    If evenOrOdd = "even" Then
        modResult = 0
    ElseIf evenOrOdd = "odd" Then
        modResult = 1
    End If
    
    Dim innerArrLength As Integer
    innerArrLength = UBound(arr, 2)
        
    
    Dim returnArrLength As Integer ' Dynamic Multidimensional Array, reDim work only last dimension
    If UBound(arr, 1) Mod 2 = 1 Then
        returnArrLength = (UBound(arr, 1) + 1) / 2
    Else
        returnArrLength = UBound(arr, 1) / 2
    End If
    
    
    Dim returnArr() As Variant
    ReDim returnArr(1 To returnArrLength, 1 To innerArrLength)
    
    Dim i, j, counter As Integer
    counter = 0
    For i = 1 To UBound(arr, 1)
        If i Mod 2 = modResult Then
            counter = counter + 1
           For j = 1 To innerArrLength
            returnArr(counter, j) = arr(i, j)
           Next j
        End If
    Next i
    
    evenOrOddIndexArrayFilter = returnArr

End Function


Private Function cropedArry(arr() As Variant, startIndex As Integer, endIndex As Integer) As Variant
'this function give croped array from source array
    
    Dim innerArrLength As Integer
    innerArrLength = UBound(arr, 2)
    
    Dim returnArrLength As Integer ' Dynamic Multidimensional Array, reDim work only last dimension
    returnArrLength = endIndex - startIndex + 1
    
    Dim returnArr() As Variant
    ReDim returnArr(1 To returnArrLength, 1 To innerArrLength)
    
    Dim i, j, counter As Integer
    counter = 0
    For i = startIndex To endIndex
            counter = counter + 1
           For j = 1 To innerArrLength
            returnArr(counter, j) = arr(i, j)
           Next j
    Next i
    
    cropedArry = returnArr

End Function


Private Function cropedArryWithStoreLastRow(inputArray As Variant, startRow As Integer, endRow As Integer) As Variant
'this function give croped array from source array with store last row number at end column
    Dim outputArray() As Variant
    ReDim outputArray(1 To endRow - startRow + 1, 1 To UBound(inputArray, 2) + 1)
    Dim i As Integer, j As Integer, k As Integer
    k = 1
    For i = startRow To endRow
        For j = 1 To UBound(inputArray, 2)
            outputArray(k, j) = inputArray(i, j)
        Next j
        outputArray(k, UBound(outputArray, 2)) = i 'store row number of original array
        k = k + 1
    Next i
    cropedArryWithStoreLastRow = outputArray
End Function



Private Function valueCounter(arr() As Variant, patternStr As String, columnIndex As Integer) As Variant ' provide source array, patternStr & column index
'this function give how many time a value exist in array, it's help to find duplicate value
    
    Dim regex As New RegExp
    regex.Global = True
    regex.pattern = patternStr
    regex.MultiLine = True
    
    Dim i, counter As Integer
    counter = 0
    For i = 1 To UBound(arr, 1)
        If regex.test(arr(i, columnIndex)) Then
            counter = counter + 1
        End If
    Next i
    
    valueCounter = counter

End Function




Private Function replaceRegExSpecialCharacterWithEscapeCharacter(regExString As Variant) As Variant  ' provide String
'this function replace regEx special character with escape character

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    regExString = Trim(regExString)
    
    regExString = Replace(regExString, "\", "\\") ' must this character at first line, otherwise replace twice
    regExString = Replace(regExString, ".", "\.")
    regExString = Replace(regExString, "^", "\^")
    regExString = Replace(regExString, "$", "\$")
    regExString = Replace(regExString, "*", "\*")
    regExString = Replace(regExString, "+", "\+")
    regExString = Replace(regExString, "-", "\-")
    regExString = Replace(regExString, "?", "\?")
    regExString = Replace(regExString, "(", "\(")
    regExString = Replace(regExString, ")", "\)")
    regExString = Replace(regExString, "[", "\[")
    regExString = Replace(regExString, "]", "\]")
    regExString = Replace(regExString, "{", "\{")
    regExString = Replace(regExString, "}", "\}")
    regExString = Replace(regExString, "|", "\|")
    regExString = Replace(regExString, vsCodeNotSupportedOrBengaliTxtDictionary("charCode151"), vsCodeNotSupportedOrBengaliTxtDictionary("charCode151WithSlash"))
    regExString = Replace(regExString, "/", "\/")
    
    
    replaceRegExSpecialCharacterWithEscapeCharacter = regExString

End Function




Private Function upClause8SpecificMushakOrBillOfEntryPreviousBalanceTransferCompare(arrUpClause8Range As Variant, sourceDataPreviousUpClause8 As Variant) As Variant
'      this function give compare result mushak & bill of entry balance from previous UP is successfully transferred or not?


    Dim arrUpClause8 As Variant
    arrUpClause8 = arrUpClause8Range.value


    Dim Result As Variant
    Dim emptyIndex As Variant
    
    Dim intialReturnArr(1 To 200, 1 To 4) As Variant

    Dim isAllResultOkArr(1 To 200, 1 To 4) As Variant


    Dim iterator As Integer
    
    For iterator = 1 To UBound(arrUpClause8, 1) - 1

    Dim sourceDataPreviousUpClause8MushakOrBillOfEntryIndex As Variant
    




    
    sourceDataPreviousUpClause8MushakOrBillOfEntryIndex = Application.Run("utilityFunction.filterMushakOrBillOfEntryArrayWithCompareQtyAndValue", arrUpClause8(iterator, 6), 6, arrUpClause8(iterator, 15), 15, arrUpClause8(iterator, 16), 16, sourceDataPreviousUpClause8)
    
    If IsArray(sourceDataPreviousUpClause8MushakOrBillOfEntryIndex) Then
    
        sourceDataPreviousUpClause8MushakOrBillOfEntryIndex = sourceDataPreviousUpClause8MushakOrBillOfEntryIndex(1, UBound(sourceDataPreviousUpClause8MushakOrBillOfEntryIndex, 2))
    
    End If
    
    If Not IsNull(sourceDataPreviousUpClause8MushakOrBillOfEntryIndex) Then
                
'        Qty.
        Dim qtyFromPreviousUp, qtyFromCurrentUp As Variant

        qtyFromPreviousUp = sourceDataPreviousUpClause8(sourceDataPreviousUpClause8MushakOrBillOfEntryIndex, 25)
        qtyFromCurrentUp = arrUpClause8(iterator, 21) + arrUpClause8(iterator, 25) ' use + balance
        
'        result = Round(qtyFromPreviousUp, 2) = Round(qtyFromCurrentUp, 2)
        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(qtyFromPreviousUp, 2), Round(qtyFromCurrentUp, 2), 0.1)
        
            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & Round(Round(qtyFromPreviousUp, 2) - Round(qtyFromCurrentUp, 2), 2)
            End If
            
        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("Q" & iterator), Result
        Application.Run "EditComment", arrUpClause8Range.Range("Q" & iterator), "Transfered Qty. Specific Check " & Result
        
        emptyIndex = Application.Run("utilityFunction.indexOf", isAllResultOkArr, "^$", 1, 1, UBound(isAllResultOkArr, 1)) ' find empty string pattern = "^$"
        

        isAllResultOkArr(emptyIndex, 1) = " "
        isAllResultOkArr(emptyIndex, 2) = " "
        isAllResultOkArr(emptyIndex, 3) = " "
        isAllResultOkArr(emptyIndex, 4) = Result
        
        
'            If Round(qtyFromPreviousUp, 2) <> Round(qtyFromCurrentUp, 2) Then
            If Not Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(qtyFromPreviousUp, 2), Round(qtyFromCurrentUp, 2), 0.1) Then
            
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                intialReturnArr(emptyIndex, 1) = "Previous Balance Compare(" & arrUpClause8(iterator, 6) & ")"
                intialReturnArr(emptyIndex, 2) = qtyFromCurrentUp & " Current UP Used + Balance (Qty.)"
                intialReturnArr(emptyIndex, 3) = qtyFromPreviousUp & " Previous UP Balance (Qty.)"
                intialReturnArr(emptyIndex, 4) = Result
            End If
            
            
'        value
        Dim valueFromPreviousUp, valueFromCurrentUp As Variant

        valueFromPreviousUp = sourceDataPreviousUpClause8(sourceDataPreviousUpClause8MushakOrBillOfEntryIndex, 26)
        valueFromCurrentUp = arrUpClause8(iterator, 22) + arrUpClause8(iterator, 26) ' use + balance

'        result = Round(valueFromPreviousUp, 2) = Round(valueFromCurrentUp, 2)
        Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(valueFromPreviousUp, 2), Round(valueFromCurrentUp, 2), 0.1)

            If Result Then
                Result = "OK"
            Else
                Result = "Mismatch = " & Round(Round(valueFromPreviousUp, 2) - Round(valueFromCurrentUp, 2), 2)
            End If

        Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("R" & iterator), Result
        Application.Run "EditComment", arrUpClause8Range.Range("R" & iterator), "Transfered Value Specific Check " & Result

        emptyIndex = Application.Run("utilityFunction.indexOf", isAllResultOkArr, "^$", 1, 1, UBound(isAllResultOkArr, 1)) ' find empty string pattern = "^$"


        isAllResultOkArr(emptyIndex, 1) = " "
        isAllResultOkArr(emptyIndex, 2) = " "
        isAllResultOkArr(emptyIndex, 3) = " "
        isAllResultOkArr(emptyIndex, 4) = Result


'                If Round(valueFromPreviousUp, 2) <> Round(valueFromCurrentUp, 2) Then
                If Not Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(valueFromPreviousUp, 2), Round(valueFromCurrentUp, 2), 0.1) Then

                    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

                    intialReturnArr(emptyIndex, 1) = "Previous Balance Compare(" & arrUpClause8(iterator, 6) & ")"
                    intialReturnArr(emptyIndex, 2) = Round(valueFromCurrentUp, 2) & " Current UP Used + Balance (Value)"
                    intialReturnArr(emptyIndex, 3) = Round(valueFromPreviousUp, 2) & " Previous UP Balance (Value)"
                    intialReturnArr(emptyIndex, 4) = Result
                End If

            

            Else
'               when mushak or bill of entry not found in previous UP then this block run

                '        Qty.
                Dim qtyFromCurrentUpNewEntry, qtyFromCurrentUpForNewEntry As Variant
                
                qtyFromCurrentUpNewEntry = arrUpClause8(iterator, 15)
                qtyFromCurrentUpForNewEntry = arrUpClause8(iterator, 21) + arrUpClause8(iterator, 25) ' use + balance
                
'                result = Round(qtyFromCurrentUpNewEntry, 2) = Round(qtyFromCurrentUpForNewEntry, 2)
                Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(qtyFromCurrentUpNewEntry, 2), Round(qtyFromCurrentUpForNewEntry, 2), 0.1)
                
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(qtyFromCurrentUpNewEntry, 2) - Round(qtyFromCurrentUpForNewEntry, 2)
                    End If
                    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("Q" & iterator), Result
                Application.Run "EditComment", arrUpClause8Range.Range("Q" & iterator), "Transfered Qty. Specific Check " & Result
                
                emptyIndex = Application.Run("utilityFunction.indexOf", isAllResultOkArr, "^$", 1, 1, UBound(isAllResultOkArr, 1)) ' find empty string pattern = "^$"
                
                
                isAllResultOkArr(emptyIndex, 1) = " "
                isAllResultOkArr(emptyIndex, 2) = " "
                isAllResultOkArr(emptyIndex, 3) = " "
                isAllResultOkArr(emptyIndex, 4) = Result
                
                
'                    If Round(qtyFromCurrentUpNewEntry, 2) <> Round(qtyFromCurrentUpForNewEntry, 2) Then
                    If Not Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(qtyFromCurrentUpNewEntry, 2), Round(qtyFromCurrentUpForNewEntry, 2), 0.1) Then
                    
                        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
                
                        intialReturnArr(emptyIndex, 1) = "New Entry(" & arrUpClause8(iterator, 6) & ")"
                        intialReturnArr(emptyIndex, 2) = Round(qtyFromCurrentUpForNewEntry, 2) & " Current UP Used + Balance (Qty.)"
                        intialReturnArr(emptyIndex, 3) = Round(qtyFromCurrentUpNewEntry, 2) & " New Entry (Qty.)"
                        intialReturnArr(emptyIndex, 4) = Result
                    End If
                    
                    
                    
                '        Value
                Dim valueFromCurrentUpNewEntry, valueFromCurrentUpForNewEntry As Variant
                
                valueFromCurrentUpNewEntry = arrUpClause8(iterator, 16)
                valueFromCurrentUpForNewEntry = arrUpClause8(iterator, 22) + arrUpClause8(iterator, 26) ' use + balance
                
'                result = Round(valueFromCurrentUpNewEntry, 2) = Round(valueFromCurrentUpForNewEntry, 2)
                Result = Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(valueFromCurrentUpNewEntry, 2), Round(valueFromCurrentUpForNewEntry, 2), 0.1)
                
                    If Result Then
                        Result = "OK"
                    Else
                        Result = "Mismatch = " & Round(valueFromCurrentUpNewEntry, 2) - Round(valueFromCurrentUpForNewEntry, 2)
                    End If
                    
                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("R" & iterator), Result
                Application.Run "EditComment", arrUpClause8Range.Range("R" & iterator), "Transfered Value Specific Check " & Result
                
                emptyIndex = Application.Run("utilityFunction.indexOf", isAllResultOkArr, "^$", 1, 1, UBound(isAllResultOkArr, 1)) ' find empty string pattern = "^$"
                
                
                isAllResultOkArr(emptyIndex, 1) = " "
                isAllResultOkArr(emptyIndex, 2) = " "
                isAllResultOkArr(emptyIndex, 3) = " "
                isAllResultOkArr(emptyIndex, 4) = Result
                
                
'                    If Round(valueFromCurrentUpNewEntry, 2) <> Round(valueFromCurrentUpForNewEntry, 2) Then
                    If Not Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(valueFromCurrentUpNewEntry, 2), Round(valueFromCurrentUpForNewEntry, 2), 0.1) Then
                    
                        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
                
                        intialReturnArr(emptyIndex, 1) = "New Entry(" & arrUpClause8(iterator, 6) & ")"
                        intialReturnArr(emptyIndex, 2) = Round(valueFromCurrentUpForNewEntry, 2) & " Current UP Used + Balance (Value)"
                        intialReturnArr(emptyIndex, 3) = Round(valueFromCurrentUpNewEntry, 2) & " New Entry (Value)"
                        intialReturnArr(emptyIndex, 4) = Result
                    End If
            End If


    Next iterator


    Dim isAllResultOkArrCropIndex As Integer
    isAllResultOkArrCropIndex = Application.Run("utilityFunction.indexOf", isAllResultOkArr, "^$", 1, 1, UBound(isAllResultOkArr, 1)) - 1 ' find empty string pattern = "^$"
    
    Dim isAllResultOkCroppedArr As Variant
    isAllResultOkCroppedArr = Application.Run("utilityFunction.cropedArry", isAllResultOkArr, 1, isAllResultOkArrCropIndex)

    
    If Application.Run("utilityFunction.isAllResultOk", isAllResultOkCroppedArr) = "OK" Then
                intialReturnArr(1, 1) = "Previous Balance Transfer(Specific check)"
                intialReturnArr(1, 2) = ""
                intialReturnArr(1, 3) = ""
                intialReturnArr(1, 4) = "OK"
    
    End If
    
    

    Dim intialReturnArrCropIndex As Integer
    intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"

    upClause8SpecificMushakOrBillOfEntryPreviousBalanceTransferCompare = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)


End Function





Private Function findNewOrPreviousBillOfEntryOrMushak(arrUpClause8Range As Range, sourceDataPreviousUpClause8 As Variant) As Variant
    'this function return two filterd array as new bill of entry or mushak and previous bill of entry or mushak
    Dim arrUpClause8 As Variant
    arrUpClause8 = arrUpClause8Range.value
    
    Dim temp As Variant
    Dim i As Integer

    For i = 1 To UBound(arrUpClause8, 1) - 1

    temp = Application.Run("utilityFunction.towDimensionalArrayFilter", sourceDataPreviousUpClause8, arrUpClause8(i, 6), 6)
    
    If IsArray(temp) Then
        
        If UBound(temp, 1) = 1 Then
            
            If Round(arrUpClause8(i, 15), 2) = Round(temp(1, 15), 2) And Round(arrUpClause8(i, 16), 2) = Round(temp(1, 16), 2) Then
                
                arrUpClause8(i, 12) = "previous"
                
            Else
                
                arrUpClause8(i, 12) = "new"
                
            End If
            
        Else
            
            Dim j As Integer
            Dim isExist As Boolean
            
            isExist = False
            
            For j = 1 To UBound(temp, 1)
                
                If Round(arrUpClause8(i, 15), 2) = Round(temp(j, 15), 2) And Round(arrUpClause8(i, 16), 2) = Round(temp(j, 16), 2) Then
                    isExist = True
                    Exit For
                End If
                
            Next j
            
            If isExist Then
                
                arrUpClause8(i, 12) = "previous"
                
            Else
                
                arrUpClause8(i, 12) = "new"
                
            End If
            
            Application.Run "EditComment", arrUpClause8Range.Range("F" & i), "Duplicate in previous UP"
            
        End If
        
    Else
        
        arrUpClause8(i, 12) = "new"
        
    End If
    
        
    Next i
    
    
    ' Store the arrays in a variant data type
    Dim resultArr(1 To 2) As Variant
    resultArr(1) = Application.Run("utilityFunction.towDimensionalArrayFilter", arrUpClause8, "previous", 12)
    resultArr(2) = Application.Run("utilityFunction.towDimensionalArrayFilter", arrUpClause8, "new", 12)
    

    findNewOrPreviousBillOfEntryOrMushak = resultArr
    
End Function




Private Function findbillOfEntryOrMushakExcludeOrTransferedFromPreviousUp(arrUpClause8Range As Range, sourceDataPreviousUpClause8 As Variant) As Variant
    'this function return two filterd array as exclude from current UP bill of entry or mushak and transfered from previous UP bill of entry or mushak
    Dim arrUpClause8 As Variant
    arrUpClause8 = arrUpClause8Range.value
    
    Dim temp As Variant
    Dim i As Integer

    For i = 1 To UBound(sourceDataPreviousUpClause8, 1) - 1

    temp = Application.Run("utilityFunction.towDimensionalArrayFilter", arrUpClause8, sourceDataPreviousUpClause8(i, 6), 6)
    
    If IsArray(temp) Then
        
        If UBound(temp, 1) = 1 Then
            
            If Round(sourceDataPreviousUpClause8(i, 15), 2) = Round(temp(1, 15), 2) And Round(sourceDataPreviousUpClause8(i, 16), 2) = Round(temp(1, 16), 2) Then
                
                sourceDataPreviousUpClause8(i, 12) = "transfer"
                
            Else
                
                sourceDataPreviousUpClause8(i, 12) = "exclude"
                
            End If
            
        Else
            
            Dim j As Integer
            Dim isExist As Boolean
            
            isExist = False
            
            For j = 1 To UBound(temp, 1)
                
                If Round(sourceDataPreviousUpClause8(i, 15), 2) = Round(temp(j, 15), 2) And Round(sourceDataPreviousUpClause8(i, 16), 2) = Round(temp(j, 16), 2) Then
                    isExist = True
                    Exit For
                End If
                
            Next j
            
            If isExist Then
                
                sourceDataPreviousUpClause8(i, 12) = "transfer"
                
            Else
                
                sourceDataPreviousUpClause8(i, 12) = "exclude"
                
            End If
            
            
        End If
        
    Else
        
        sourceDataPreviousUpClause8(i, 12) = "exclude"
        
    End If
    
        
    Next i
    
    
    ' Store the arrays in a variant data type
    Dim resultArr(1 To 2) As Variant
    resultArr(1) = Application.Run("utilityFunction.towDimensionalArrayFilter", sourceDataPreviousUpClause8, "transfer", 12)
    resultArr(2) = Application.Run("utilityFunction.towDimensionalArrayFilter", sourceDataPreviousUpClause8, "exclude", 12)
    

    findbillOfEntryOrMushakExcludeOrTransferedFromPreviousUp = resultArr
    
End Function




Private Function indexOfMushakOrBillOfEntryWithCompareQtyAndValue(mushakOrBillOfEntry As Variant, mushakOrBillOfEntryFindColumn As Integer, qty As Variant, qtyColumn As Integer, value As Variant, valueColumn As Integer, mushakOrBillOfEntryArrForFind As Variant) As Variant
    'this function return index of mushak or bill of entry with compare Qty. & Value
    
    Dim regex As New RegExp
    regex.Global = True
    regex.MultiLine = True
    
    Dim temp As Variant
    Dim index As Integer
    
    temp = Application.Run("utilityFunction.towDimensionalArrayFilter", mushakOrBillOfEntryArrForFind, mushakOrBillOfEntry, mushakOrBillOfEntryFindColumn)
    
    If IsArray(temp) Then '1
        
        If UBound(temp, 1) = 1 Then '2
            
            index = temp(1, UBound(temp, 2))
            
        Else '2
            
            regex.pattern = "\d+"
            
            Set qty = regex.Execute(qty)
            
            qty = qty.Item(0)
            
            temp = Application.Run("utilityFunction.towDimensionalArrayFilter", temp, qty, qtyColumn)
            
            If IsArray(temp) Then '3
                
                If UBound(temp, 1) = 1 Then '4
                    
                    index = temp(1, UBound(temp, 2) - 1)
                    
                Else '4
                    
                    regex.pattern = "\d+"
                    
                    Set value = regex.Execute(value)
                    
                    value = value.Item(0)
                    
                    temp = Application.Run("utilityFunction.towDimensionalArrayFilter", temp, value, valueColumn)
                    
                    If IsArray(temp) Then '5
                        
                        index = temp(1, UBound(temp, 2) - 2)
                        
                    Else '5
                        
                        indexOfMushakOrBillOfEntryWithCompareQtyAndValue = Null
                        
                        Exit Function
                        
                    End If '5
                    
                End If '4
                
            Else '3
                
                indexOfMushakOrBillOfEntryWithCompareQtyAndValue = Null
                
                Exit Function
                
            End If '3
            
        End If '2
        
    Else '1
        
        indexOfMushakOrBillOfEntryWithCompareQtyAndValue = Null
        
        Exit Function
        
    End If '1
    
    indexOfMushakOrBillOfEntryWithCompareQtyAndValue = index
    
End Function







Private Function mergeArry(mainArr() As Variant, addingArr() As Variant, columnForStartingRowFinding As Integer) As Variant ' provide main arr, adding arr & column number of main arr to find the starting row
'this function give merge array from two source array
    
    Dim innerArrLengthAddingArr, innerArrLengthMainArr As Integer
    innerArrLengthAddingArr = UBound(addingArr, 2)
    innerArrLengthMainArr = UBound(mainArr, 2)

    Dim outerArrLengthAddingArr, outerArrLengthMainArr As Integer
    outerArrLengthAddingArr = UBound(addingArr, 1)
    outerArrLengthMainArr = UBound(mainArr, 1)
    
    Dim emptyIndex As Integer
    
    emptyIndex = Application.Run("utilityFunction.indexOf", mainArr, "^$", 1, 1, UBound(mainArr, 1)) ' find empty string pattern = "^$"
    
    Dim counter As Integer
    
    counter = Application.Run("utilityFunction.indexOf", mainArr, "^$", columnForStartingRowFinding, 1, UBound(mainArr, 1)) - 1  ' find empty string pattern = "^$"
    
    

    If innerArrLengthAddingArr <= innerArrLengthMainArr And outerArrLengthMainArr - counter > outerArrLengthAddingArr Then
    
        Dim i, j As Integer
        For i = 1 To UBound(addingArr, 1)
                counter = counter + 1
            For j = 1 To innerArrLengthAddingArr
                mainArr(counter, j) = addingArr(i, j)
            Next j
        Next i
    
        mergeArry = mainArr
        Exit Function
    Else
    
        Dim k As Integer
            For k = 1 To innerArrLengthAddingArr
                mainArr(emptyIndex, k) = "merge missing"
            Next k
    
        mergeArry = mainArr
         
    End If

End Function



Private Function addSpecificStringTo2DArrayOnLastColumn(arr As Variant, str As Variant) As Variant
    ' this function receive an array and specific string, and return new array with extra column, extra column contain specific string
    If IsArray(arr) Then

        Dim numRows As Long, numCols As Long
        numRows = UBound(arr, 1) - LBound(arr, 1) + 1
        numCols = UBound(arr, 2) - LBound(arr, 2) + 1
        
        Dim outputArr() As Variant
        ReDim outputArr(1 To numRows, 1 To numCols + 1)
        
        Dim i As Long, j As Long
        For i = 1 To numRows
            For j = 1 To numCols
                outputArr(i, j) = arr(i, j)
            Next j
            outputArr(i, numCols + 1) = str
        Next i

    
    Else

        addSpecificStringTo2DArrayOnLastColumn = Null
    
        Exit Function
    
    End If
    
    addSpecificStringTo2DArrayOnLastColumn = outputArr
    
End Function






Private Function mLcUdExpIpCompareWithSource(mLcUdExpIpFromUpClause As Variant, mLcUdExpIpFromSourceData As Variant, mLcUdExpIpDateFromSourceData As Variant, whatCheck As Variant) As Variant
'      this function give compare result M.LC UD EXP IP of UP clause & source data

    Dim regex As New RegExp
    regex.Global = True
    
    
    Dim temp As Variant
    Dim emptyIndex As Variant
    
    Dim intialReturnArr(1 To 50, 1 To 4) As Variant
    
    Dim regExReturnedMLcUdExpIpFromSourceDataObject, regExReturnedMLcUdExpIpDateFromSourceDataObject As Variant
    
    regex.pattern = ".+"
    Set regExReturnedMLcUdExpIpFromSourceDataObject = regex.Execute(mLcUdExpIpFromSourceData)
    
    Set regExReturnedMLcUdExpIpDateFromSourceDataObject = regex.Execute(mLcUdExpIpDateFromSourceData)
    
    Dim iterator As Integer
    
    For iterator = 0 To regExReturnedMLcUdExpIpFromSourceDataObject.Count - 1
    
        regex.pattern = Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", regExReturnedMLcUdExpIpFromSourceDataObject.Item(iterator))
        temp = regex.test(mLcUdExpIpFromUpClause)
        
        If temp Then
            
            Dim tempMLcUdExpIpFromUpClauseWithDate As Variant
            
            regex.pattern = Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", regExReturnedMLcUdExpIpFromSourceDataObject.Item(iterator)) & "[\s\S]{0,10}\d{2}\/\d{2}\/\d{4}"
            
            Set tempMLcUdExpIpFromUpClauseWithDate = regex.Execute(mLcUdExpIpFromUpClause)
            
            regex.pattern = "Dt\..*\d{2}\/\d{2}\/\d{4}$|\d{2}\/\d{2}\/\d{4}$|\n"
            
            temp = Trim(regex.Replace(tempMLcUdExpIpFromUpClauseWithDate(0), ""))
            
            
            temp = temp = regExReturnedMLcUdExpIpFromSourceDataObject.Item(iterator)
            
            If temp Then
            
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = whatCheck
                intialReturnArr(emptyIndex, 2) = Trim(regex.Replace(tempMLcUdExpIpFromUpClauseWithDate(0), ""))
                intialReturnArr(emptyIndex, 3) = regExReturnedMLcUdExpIpFromSourceDataObject.Item(iterator)
                intialReturnArr(emptyIndex, 4) = "OK"
                
            Else
                emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
    
                intialReturnArr(emptyIndex, 1) = whatCheck
                intialReturnArr(emptyIndex, 2) = Trim(regex.Replace(tempMLcUdExpIpFromUpClauseWithDate(0), ""))
                intialReturnArr(emptyIndex, 3) = regExReturnedMLcUdExpIpFromSourceDataObject.Item(iterator)
                intialReturnArr(emptyIndex, 4) = "Mismatch"
            End If
            
            
            regex.pattern = Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", regExReturnedMLcUdExpIpDateFromSourceDataObject.Item(iterator)) & "$"
            
            temp = regex.test(tempMLcUdExpIpFromUpClauseWithDate(0))
            
                If temp Then
                
                    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
                
                    intialReturnArr(emptyIndex, 1) = whatCheck & " Date"
                    intialReturnArr(emptyIndex, 2) = CDate(Right(tempMLcUdExpIpFromUpClauseWithDate(0), 10))
                    intialReturnArr(emptyIndex, 3) = CDate(regExReturnedMLcUdExpIpDateFromSourceDataObject.Item(iterator))
                    intialReturnArr(emptyIndex, 4) = "OK"
                
                Else
        
                    emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
                
                    intialReturnArr(emptyIndex, 1) = whatCheck & " Date"
                    intialReturnArr(emptyIndex, 2) = CDate(Right(tempMLcUdExpIpFromUpClauseWithDate(0), 10))
                    intialReturnArr(emptyIndex, 3) = CDate(regExReturnedMLcUdExpIpDateFromSourceDataObject.Item(iterator))
                    intialReturnArr(emptyIndex, 4) = "Mismatch"
                
                
                End If
        
        Else
    
            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"
        
            intialReturnArr(emptyIndex, 1) = whatCheck
            intialReturnArr(emptyIndex, 2) = ""
            intialReturnArr(emptyIndex, 3) = regExReturnedMLcUdExpIpFromSourceDataObject.Item(iterator)
            intialReturnArr(emptyIndex, 4) = "Not Found"
        
        End If
    
    
    Next iterator
    

    Dim intialReturnArrCropIndex As Integer
    intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"

    mLcUdExpIpCompareWithSource = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)


End Function



Private Function expOrIpExtractorFromSourceData(expAndIpString As Variant, whatExtract As String) As Variant ' provide exp + ip string, extract criteria ("exp" or "ip")
'this function give extracted string from source string, based on criteria
    Dim patternStr As Variant
    Dim whatReplace As String
    If whatExtract = "exp" Then
        patternStr = "EXP\:.+"
        whatReplace = "EXP:"
    ElseIf whatExtract = "ip" Then
        patternStr = "IP\:.+"
        whatReplace = "IP:"
    End If
    
    Dim regex As New RegExp
    regex.Global = True
    regex.MultiLine = True
    regex.pattern = patternStr
    
    Dim regExReturnedObjectExpOrIp, regExReturnedExtractedExpOrIp As Variant

    Set regExReturnedObjectExpOrIp = regex.Execute(expAndIpString)

    Dim iterator As Integer

    For iterator = 0 To regExReturnedObjectExpOrIp.Count - 1
            
            regExReturnedExtractedExpOrIp = regExReturnedExtractedExpOrIp & " " & Trim(Replace(regExReturnedObjectExpOrIp.Item(iterator), whatReplace, ""))
        
    Next iterator
        
    
    expOrIpExtractorFromSourceData = Replace(Trim(regExReturnedExtractedExpOrIp), " ", Chr(10))

End Function




Private Function expOrIpDateExtractorFromSourceDate(expOrIpFromSourceData As Variant, expOrIpDateFromSourceData As Variant, extractedExpOrIpFromSourceDate As Variant) As Variant
'      this function give extracted EXP or IP date from source data

    Dim regex As New RegExp
    regex.Global = True
    regex.MultiLine = True
    
    Dim temp As Variant
    Dim Result As Variant
    
    Dim regExReturnedExpOrIpFromSourceDataObject, regExReturnedExpOrIpDateFromSourceDataObject, regExReturnedExtractedExpOrIpFromSourceDataObject As Variant
    
    regex.pattern = ".+"

    Set regExReturnedExpOrIpFromSourceDataObject = regex.Execute(expOrIpFromSourceData)

    Set regExReturnedExpOrIpDateFromSourceDataObject = regex.Execute(expOrIpDateFromSourceData)
    
    Set regExReturnedExtractedExpOrIpFromSourceDataObject = regex.Execute(extractedExpOrIpFromSourceDate)
    
    Dim iterator As Integer
    
    For iterator = 0 To regExReturnedExtractedExpOrIpFromSourceDataObject.Count - 1
    
        regex.pattern = Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", regExReturnedExtractedExpOrIpFromSourceDataObject.Item(iterator))

        Dim i As Integer

        For i = 0 To regExReturnedExpOrIpFromSourceDataObject.Count - 1

            temp = regex.test(regExReturnedExpOrIpFromSourceDataObject.Item(i))

            If temp Then

                Result = Result & " " & regExReturnedExpOrIpDateFromSourceDataObject.Item(i)

            End If
        
        Next i
    
    
    Next iterator

    Result = Replace(Trim(Result), " ", Chr(10))
    

    expOrIpDateExtractorFromSourceDate = Result

    

End Function




Private Function isAllResultOk(resultArr As Variant) As Variant
'      this function returned ("OK"), if all result are ok

    Dim iterator As Integer
    
    For iterator = 1 To UBound(resultArr, 1)
    
       If resultArr(iterator, 4) <> "OK" Then

            isAllResultOk = "Mismatch"

            Exit Function
            
       End If
    
    Next iterator
    

    isAllResultOk = "OK"


End Function







Private Function upClause8MushakOrBillOfEntryCompare(arrUpClause8Range As Variant, upClause8mushakOrBillOfEntryArr As Variant, sourceData As Variant, sourceDataBillOfEntryOrMushakColumn As Integer, sourceDataLcColumn As Integer, sourceDataQtyColumn As Integer, sourceDataValueColumn As Integer, isMushakOrBillOfEntry As Variant, classification As Variant) As Variant ' provide UP clause8 range, UP clase8 classified arr, source data arr, mushak or bill of entry ("Mushak"/"Bill of Entry"), Bill of entry or Mushak Column, Bill of entry or Mushak LC Column, Qty. Column, Value Column & classification (local yarn/ import yarn/ dyes/ local chemical/ import chemical)
    'this function compare mushak & bill of entry with source data push and result return array and also mark UP sheet

       Dim regex As New RegExp
       regex.Global = True
       regex.MultiLine = True

        Dim upClause8Arr As Variant
        Dim indexOfRangeMarking As Integer

        upClause8Arr = arrUpClause8Range.value

        Dim intialReturnArr(1 To 200, 1 To 4) As Variant

        Dim emptyIndex As Variant
        Dim Result As Variant

        If IsArray(upClause8mushakOrBillOfEntryArr) Then ' if mushak or bill of entry exist

        Dim specificMushakOrBillOfEntryInformationFromSourceData As Variant

        Dim MushakOrBillOfEntryIterator As Integer

        For MushakOrBillOfEntryIterator = 1 To UBound(upClause8mushakOrBillOfEntryArr, 1)

        Dim patternStr As String

        Dim clause8MushakOrBillOfEntryAndDate, clause8MushakOrBillOfEntryLcAndDate, clause8MushakOrBillOfEntryQty, clause8MushakOrBillOfEntryValue As Variant
        Dim sourceDataMushakOrBillOfEntryAndDate, sourceDataMushakOrBillOfEntryLcAndDate, sourceDataMushakOrBillOfEntryQty, sourceDataMushakOrBillOfEntryValue As Variant


        clause8MushakOrBillOfEntryAndDate = upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 6)

        clause8MushakOrBillOfEntryLcAndDate = upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 2)

        clause8MushakOrBillOfEntryQty = upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 15)

        clause8MushakOrBillOfEntryValue = upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 16)

            'filter by LC
        specificMushakOrBillOfEntryInformationFromSourceData = Application.Run("utilityFunction.towDimensionalArrayFilter", sourceData, _ 
            "^" & Application.Run("utilityFunction.replaceRegExSpecialCharacterWithEscapeCharacter", clause8MushakOrBillOfEntryLcAndDate), 4)

        If IsArray(specificMushakOrBillOfEntryInformationFromSourceData) Then

            specificMushakOrBillOfEntryInformationFromSourceData = Application.Run("utilityFunction.filterMushakOrBillOfEntryArrayWithCompareQtyAndValue", _
                clause8MushakOrBillOfEntryAndDate, 3, clause8MushakOrBillOfEntryQty, sourceDataQtyColumn, clause8MushakOrBillOfEntryValue, sourceDataValueColumn, _
                    specificMushakOrBillOfEntryInformationFromSourceData)

        End If

        emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

        Dim mushakOrBillOfEntryOnly As Variant

        regex.pattern = ".+"
        Set mushakOrBillOfEntryOnly = regex.Execute(clause8MushakOrBillOfEntryAndDate)
        mushakOrBillOfEntryOnly = mushakOrBillOfEntryOnly.Item(0)


        intialReturnArr(emptyIndex, 1) = " "
        intialReturnArr(emptyIndex, 2) = isMushakOrBillOfEntry & "(" & classification & ") " & MushakOrBillOfEntryIterator & ") " & mushakOrBillOfEntryOnly
        intialReturnArr(emptyIndex, 3) = ""
        intialReturnArr(emptyIndex, 4) = ""


        If IsNull(specificMushakOrBillOfEntryInformationFromSourceData) Then


            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = isMushakOrBillOfEntry
            intialReturnArr(emptyIndex, 2) = clause8MushakOrBillOfEntryAndDate
            intialReturnArr(emptyIndex, 3) = ""
            intialReturnArr(emptyIndex, 4) = "Not found in import performance"

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("f" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), "Mismatch"
            Application.Run "EditComment", arrUpClause8Range.Range("f" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), "Not found in import performance"

        ElseIf UBound(specificMushakOrBillOfEntryInformationFromSourceData, 1) = 1 Then

            regex.pattern = clause8MushakOrBillOfEntryAndDate

            Result = regex.test(specificMushakOrBillOfEntryInformationFromSourceData(1, sourceDataBillOfEntryOrMushakColumn))


                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch"
                End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("f" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), Result
            Application.Run "EditComment", arrUpClause8Range.Range("f" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), "Compare in import performance " & Result
            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = isMushakOrBillOfEntry & " & Date"
            intialReturnArr(emptyIndex, 2) = clause8MushakOrBillOfEntryAndDate
            intialReturnArr(emptyIndex, 3) = " "
            intialReturnArr(emptyIndex, 4) = Result



            regex.pattern = clause8MushakOrBillOfEntryLcAndDate

            Result = regex.test(specificMushakOrBillOfEntryInformationFromSourceData(1, sourceDataLcColumn))


                If Result Then
                    Result = "OK"
                Else
                    Result = "Mismatch"
                End If

                Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("b" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), Result
                Application.Run "EditComment", arrUpClause8Range.Range("b" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), "Compare in import performance " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "LC & Date"
            intialReturnArr(emptyIndex, 2) = clause8MushakOrBillOfEntryLcAndDate
            intialReturnArr(emptyIndex, 3) = specificMushakOrBillOfEntryInformationFromSourceData(1, sourceDataLcColumn)
            intialReturnArr(emptyIndex, 4) = Result



            sourceDataMushakOrBillOfEntryQty = specificMushakOrBillOfEntryInformationFromSourceData(1, sourceDataQtyColumn)

            Result = Round(clause8MushakOrBillOfEntryQty, 2) = Round(sourceDataMushakOrBillOfEntryQty, 2)


            If Result Then
                    Result = "OK"
            Else
                    Result = "Mismatch = " & Round(clause8MushakOrBillOfEntryQty, 2) - Round(sourceDataMushakOrBillOfEntryQty, 2)
            End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("o" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), Result
            Application.Run "EditComment", arrUpClause8Range.Range("o" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), "Compare in import performance " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "Qty."
            intialReturnArr(emptyIndex, 2) = clause8MushakOrBillOfEntryQty
            intialReturnArr(emptyIndex, 3) = sourceDataMushakOrBillOfEntryQty
            intialReturnArr(emptyIndex, 4) = Result


            sourceDataMushakOrBillOfEntryValue = specificMushakOrBillOfEntryInformationFromSourceData(1, sourceDataValueColumn)


            Result = Round(clause8MushakOrBillOfEntryValue, 2) = Round(sourceDataMushakOrBillOfEntryValue, 2)


            If Result Then
                    Result = "OK"
            Else
                    Result = "Mismatch = " & Round(clause8MushakOrBillOfEntryValue, 2) - Round(sourceDataMushakOrBillOfEntryValue, 2)
            End If

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("p" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), Result
            Application.Run "EditComment", arrUpClause8Range.Range("p" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), "Compare in import performance " & Result

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = "Value"
            intialReturnArr(emptyIndex, 2) = clause8MushakOrBillOfEntryValue
            intialReturnArr(emptyIndex, 3) = sourceDataMushakOrBillOfEntryValue
            intialReturnArr(emptyIndex, 4) = Result



        ElseIf UBound(specificMushakOrBillOfEntryInformationFromSourceData, 1) <> 1 Then

            Application.Run "utilityFunction.errorMarkingForValue", arrUpClause8Range.Range("f" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), "Mismatch"
            Application.Run "EditComment", arrUpClause8Range.Range("f" & upClause8mushakOrBillOfEntryArr(MushakOrBillOfEntryIterator, 27)), "Duplicate in import performance"

            emptyIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) ' find empty string pattern = "^$"

            intialReturnArr(emptyIndex, 1) = isMushakOrBillOfEntry
            intialReturnArr(emptyIndex, 2) = MushakOrBillOfEntryIterator & ") " & clause8MushakOrBillOfEntryAndDate
            intialReturnArr(emptyIndex, 3) = ""
            intialReturnArr(emptyIndex, 4) = "Duplicate in import performance"


        End If

        Next MushakOrBillOfEntryIterator


        End If


        Dim intialReturnArrCropIndex As Integer
        intialReturnArrCropIndex = Application.Run("utilityFunction.indexOf", intialReturnArr, "^$", 1, 1, UBound(intialReturnArr, 1)) - 1 ' find empty string pattern = "^$"

        upClause8MushakOrBillOfEntryCompare = Application.Run("utilityFunction.cropedArry", intialReturnArr, 1, intialReturnArrCropIndex)





End Function








Private Function errorMarkingForValue(errRange As Range, Result As Variant) As Variant
'    this function mark provided range based on result
    
    
    If Result = "OK" Then
        errRange.Interior.Color = RGB(0, 125, 0)
    Else
        errRange.Interior.Color = RGB(255, 0, 0)
    End If
    
    
End Function



Private Function EditComment(rng As Range, Result As Variant) As Variant

        Dim OldComment As Variant
        Dim NewComment As Variant
        

    If rng.Comment Is Nothing Then   'check if the cell has no comment

        rng.AddComment
        
        Dim cmt As Comment
        Set cmt = rng.Comment ' Range to the cell with the comment want to resize
    
        With cmt.Shape
            .Height = 100 ' Change to the desired height in points
            .Width = 500 ' Change to the desired width in points
        End With
        
    End If
    

        
    If Not rng.Comment Is Nothing Then   'check if the cell has a comment
            
        OldComment = rng.Comment.Text
        NewComment = OldComment & Result & Chr(10)

        rng.Comment.Text NewComment 'change the comment text to the desired text

    End If
    


End Function



Private Function resultSheetFormating(formatRange As Range, InteriorColor As String, fontColor As String, borderColor As String) As Variant
'    this function format result sheet
    
    With formatRange
        .Interior.Color = InteriorColor
        .Font.Color = fontColor
    End With
    
    With formatRange.Range("a1:d2")
        .Font.Size = 13
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80)
    End With
    
    
    formatRange.Borders(xlDiagonalDown).LineStyle = xlNone
    formatRange.Borders(xlDiagonalUp).LineStyle = xlNone
    With formatRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = borderColor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = borderColor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = borderColor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = borderColor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = borderColor
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With formatRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = borderColor
        .TintAndShade = 0
        .Weight = xlThin
    End With

    With Cells
    .Columns.AutoFit
    .Rows.AutoFit
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    End With
    
    Columns("A:A").ColumnWidth = 38
    Columns("B:C").ColumnWidth = 95
    Columns("D:D").ColumnWidth = 32
    
End Function


    


    





'####################################this function work fine but one issue is always not check till value so, not perfect, try to solve it bellow

'Private Function filterMushakOrBillOfEntryArrayWithCompareQtyAndValue(mushakOrBillOfEntry As Variant, mushakOrBillOfEntryFindColumn As Integer, qty As Variant, qtyFindColumn As Integer, value As Variant, valueFindColumn As Integer, mushakOrBillOfEntryArrForFind As Variant) As Variant
'    'this function return array of mushak or bill of entry with compare Qty. & Value
'
'    Dim regEx As New RegExp
'    regEx.Global = True
'    regEx.MultiLine = True
'
'    Dim temp As Variant
'    Dim qtyArray, valueArray As Variant
'    Dim i As Long, j As Long
'
'    Dim returnArray As Variant
'
'    temp = Application.Run("utilityFunction.towDimensionalArrayFilter", mushakOrBillOfEntryArrForFind, mushakOrBillOfEntry, mushakOrBillOfEntryFindColumn)
'
'    If IsArray(temp) Then '1
'
'        If UBound(temp, 1) = 1 Then '2
'
'            returnArray = temp
'
'        Else '2
'
'            regEx.Pattern = "\d+"
'
'            Set qty = regEx.Execute(qty)
'
'            qty = qty.Item(0)
'
'            temp = Application.Run("utilityFunction.towDimensionalArrayFilter", temp, qty, qtyFindColumn)
'
'            If IsArray(temp) Then '3
'
'                If UBound(temp, 1) = 1 Then '4
'
'                    ' Create new qty. array
'                    ReDim qtyArray(1 To 1, 1 To UBound(temp, 2) - 1)
'                    ' Copy data from temp array to qty. array
'                    For i = 1 To 1
'                        For j = 1 To UBound(temp, 2) - 1
'                            qtyArray(i, j) = temp(i, j)
'                        Next j
'                    Next i
'
'                    returnArray = qtyArray
'
'                Else '4
'
'                    regEx.Pattern = "\d+"
'
'                    Set value = regEx.Execute(value)
'
'                    value = value.Item(0)
'
'                    temp = Application.Run("utilityFunction.towDimensionalArrayFilter", temp, value, valueFindColumn)
'
'                    If IsArray(temp) Then '5
'
'                        ' Create new qty. array
'                        ReDim valueArray(1 To UBound(temp, 1), 1 To UBound(temp, 2) - 2)
'                        ' Copy data from temp array to qty. array
'                        For i = 1 To UBound(temp, 1)
'                            For j = 1 To UBound(temp, 2) - 2
'                                valueArray(i, j) = temp(i, j)
'                            Next j
'                        Next i
'
'                    returnArray = valueArray
'
'
'                    Else '5
'
'                        filterMushakOrBillOfEntryArrayWithCompareQtyAndValue = Null
'
'                        Exit Function
'
'                    End If '5
'
'                End If '4
'
'            Else '3
'
'                filterMushakOrBillOfEntryArrayWithCompareQtyAndValue = Null
'
'                Exit Function
'
'            End If '3
'
'        End If '2
'
'    Else '1
'
'        filterMushakOrBillOfEntryArrayWithCompareQtyAndValue = Null
'
'        Exit Function
'
'    End If '1
'
'    filterMushakOrBillOfEntryArrayWithCompareQtyAndValue = returnArray
'
'End Function


Private Function filterMushakOrBillOfEntryArrayWithCompareQtyAndValue(mushakOrBillOfEntry As Variant, mushakOrBillOfEntryFindColumn As Integer, qty As Variant, qtyFindColumn As Integer, value As Variant, valueFindColumn As Integer, mushakOrBillOfEntryArrForFind As Variant) As Variant
    'this function return array of mushak or bill of entry with compare Qty. & Value

    Dim regex As New RegExp
    regex.Global = True
    regex.MultiLine = True

    Dim temp As Variant
    Dim qtyArray, valueArray As Variant
    Dim i As Long, j As Long

    Dim returnArray As Variant

    temp = Application.Run("utilityFunction.towDimensionalArrayFilter", mushakOrBillOfEntryArrForFind, mushakOrBillOfEntry, mushakOrBillOfEntryFindColumn)

    If IsArray(temp) Then
        ' code Qty.
        regex.pattern = "\d+"

        Set qty = regex.Execute(qty)

        qty = qty.Item(0)

        temp = Application.Run("utilityFunction.towDimensionalArrayFilter", temp, qty, qtyFindColumn)

        If IsArray(temp) Then
            ' code value
            regex.pattern = "\d+"

            Set value = regex.Execute(value)

            value = value.Item(0)

            temp = Application.Run("utilityFunction.towDimensionalArrayFilter", temp, value, valueFindColumn)

            If IsArray(temp) Then
                'loop
                ' Create new value array
                ReDim valueArray(1 To UBound(temp, 1), 1 To UBound(temp, 2) - 2)
                ' Copy data from temp array to value array
                For i = 1 To UBound(temp, 1)
                    For j = 1 To UBound(temp, 2) - 2
                        valueArray(i, j) = temp(i, j)
                    Next j
                Next i

                returnArray = valueArray
            Else
                returnArray = Null
            End If
        Else
            returnArray = Null
        End If
    Else
        returnArray = Null
    End If

    filterMushakOrBillOfEntryArrayWithCompareQtyAndValue = returnArray

End Function





Private Function concatSpecificColumnString(arr As Variant, columnIndex As Integer) As Variant
    'this function give concatSrt of specific column
    
    If IsArray(arr) Then
        Dim concatSrt As Variant
        concatSrt = ""
        Dim i As Long
        For i = 1 To UBound(arr, 1)
            
            If i = 1 Then
                concatSrt = arr(i, columnIndex)
            Else
                concatSrt = concatSrt & vbNewLine & arr(i, columnIndex)
            End If
            
        Next i
        
    Else
        concatSpecificColumnString = Null
        
        Exit Function
        
    End If
    
    concatSpecificColumnString = concatSrt
    
End Function





    Private Function addSpecificColumnValueOneRowUpTo2DArrayOnLastColumn(arr As Variant, columnNo As Integer) As Variant
        ' this function receive an array and specific column no, and return new array with extra column, extra column contain received column value on one row up
        If IsArray(arr) And UBound(arr, 1) > 1 Then
    
            Dim numRows As Long, numCols As Long
            numRows = UBound(arr, 1)
            numCols = UBound(arr, 2)
            
            Dim outputArr() As Variant
            ReDim outputArr(1 To numRows, 1 To numCols + 1)
            
            Dim i As Long, j As Long
            For i = 1 To numRows - 1
            
            
                For j = 1 To numCols
                    outputArr(i, j) = arr(i, j)
                Next j
                outputArr(i, numCols + 1) = arr(i + 1, columnNo)
            Next i
        
        
        Else
    
            addSpecificColumnValueOneRowUpTo2DArrayOnLastColumn = Null
            
            Exit Function
        
        End If
    
    
        addSpecificColumnValueOneRowUpTo2DArrayOnLastColumn = outputArr
        
    End Function
    
    





        Private Function mergeTow2DArrays(arr1 As Variant, arr2 As Variant) As Variant
            ' this function received two 2D array and merged them into one and return
            
            If IsArray(arr1) And IsArray(arr2) Then
                
                If UBound(arr1, 2) = UBound(arr2, 2) Then
                    
                    Dim numRows As Long, numCols As Long
                    Dim mergedArray As Variant
                    Dim i As Long, j As Long
                    
                    ' Determine number of rows and columns in merged array
                    numRows = UBound(arr1, 1) + UBound(arr2, 1)
                    numCols = UBound(arr1, 2)
                    
                    ' Create new merged array
                    ReDim mergedArray(1 To numRows, 1 To numCols)
                    
                    ' Copy data from first array to merged array
                    For i = 1 To UBound(arr1, 1)
                        For j = 1 To UBound(arr1, 2)
                            mergedArray(i, j) = arr1(i, j)
                        Next j
                    Next i
                    
                    ' Copy data from second array to merged array
                    For i = 1 To UBound(arr2, 1)
                        For j = 1 To UBound(arr2, 2)
                            mergedArray(i + UBound(arr1, 1), j) = arr2(i, j)
                        Next j
                    Next i
                    
                Else
                    mergeTow2DArrays = Null
                    
                    Exit Function
                    
                End If
                
            Else
                
                mergeTow2DArrays = Null
                
                Exit Function
                
            End If
            
            ' Return merged array
            mergeTow2DArrays = mergedArray
        End Function
        

            



Private Function upClause8ReportSumAsclassification(arrRange As Range, arr As Variant, classification As Variant, filterColumn As Integer, sumColumn As Integer) As Variant
'    only for audit report
    
    
  
    
    Dim filterArr As Variant
    Dim sum As Variant

    filterArr = Application.Run("utilityFunction.towDimensionalArrayFilter", arr, classification, filterColumn)
    
    
    
    If IsArray(filterArr) Then
    
        sum = Application.Run("utilityFunction.sumArrColumn", filterArr, sumColumn)
        
        
        
        Dim i As Integer
        
        For i = 1 To UBound(filterArr, 1)
        
            Application.Run "utilityFunction.errorMarkingForValue", arrRange.Range("A" & filterArr(i, 24) * 2 - 1), "OK"
            
        Next i
        
    Else
    
        sum = 0
        
    End If
    
    upClause8ReportSumAsclassification = sum
    
End Function








Private Function upClause8Report(arrUpClause8Range As Range, upName As Variant) As Variant
        '    only for audit report
            Dim arrUpClause8 As Variant
            
            arrUpClause8 = arrUpClause8Range.value
            
            Dim upClause8OddFiltered As Variant
            
            upClause8OddFiltered = Application.Run("utilityFunction.evenOrOddIndexArrayFilter", arrUpClause8, "odd", True)
            
            Dim returnArr(1 To 1, 1 To 27)
            
    
            Dim patternStr As Variant
    
           
            Dim yarnUsedSum As Variant
    
            patternStr = "Yarn|YARN"
            
            yarnUsedSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18) ' only for sum check
            
            
            
            Dim wettingAgentSum As Variant
            patternStr = "Wetting Agent|Mercerizing Agent|Surface-Active Preparations"
            
            wettingAgentSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 1) = wettingAgentSum
            
            
            Dim modifiedStarchSum As Variant
            patternStr = "Modified Starch"
            
            modifiedStarchSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 2) = modifiedStarchSum
            
            
            Dim causticSodaSum As Variant
            patternStr = "Caustic Soda"
            
            causticSodaSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 3) = causticSodaSum
            
            
            Dim sulphuricAcidSum As Variant
            patternStr = "Sulphuric Acid"
            
            sulphuricAcidSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 4) = sulphuricAcidSum
            
            
            Dim reducingAgentSum As Variant
            patternStr = "Reducing Agent"
            
            reducingAgentSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 5) = reducingAgentSum
            
            
            Dim softenerSum As Variant
            patternStr = "Softening Agent|Softning Agent"
            
            softenerSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 6) = softenerSum
            
            
            Dim binderSum As Variant
            patternStr = "Binder"
            
            binderSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 7) = binderSum
            
            
            Dim sequesteringAgentSum As Variant
            patternStr = "Sequestering"
            
            sequesteringAgentSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 8) = sequesteringAgentSum
            
            
            Dim sodiumHydroSulphateSum As Variant
            patternStr = "Sodium Sulphides|Sodium Hydrosulphite"
            
            sodiumHydroSulphateSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 9) = sodiumHydroSulphateSum
            
            
            Dim waxSum As Variant
            patternStr = "Wax"
            
            waxSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 10) = waxSum
            
            
            Dim aceticAcidGreenAcidSum As Variant
            patternStr = "Acetic Acid"
            
            aceticAcidGreenAcidSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 11) = aceticAcidGreenAcidSum
            
            
            Dim PVASum As Variant
            patternStr = "PVA"
            
            PVASum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 12) = PVASum
            
            
            Dim enzymeSum As Variant
            patternStr = "Enzyme"
            
            enzymeSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 13) = enzymeSum
            
            
            Dim fixingAgentSum As Variant
            patternStr = "Fixing Agent|Sizing Agent"
            
            fixingAgentSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 14) = fixingAgentSum
            
            
            Dim dispersingAgentSum As Variant
            patternStr = "Dispersing Agent"
            
            dispersingAgentSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 15) = dispersingAgentSum
            
            ' Column 16 no import
            
            Dim waterDecoloringAgentSum As Variant
            patternStr = "Water Decoloring Agent|DE Coloring Agent|DE-Coloring Agent"
            
            waterDecoloringAgentSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 17) = waterDecoloringAgentSum
            
            
            Dim hydrogenPeroxideSum As Variant
            patternStr = "Hydrogen Peroxide"
            
            hydrogenPeroxideSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 18) = hydrogenPeroxideSum
            
            
            Dim stabilizingAgentSum As Variant
            patternStr = "Stabilizing Agent|Estabilizador FE"
            
            stabilizingAgentSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 19) = stabilizingAgentSum
            
            
            Dim detergentSum As Variant
            patternStr = "Detergent"
            
            detergentSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 20) = detergentSum
            
            
            ' Column 21, 22 no import
            
            
            Dim vatDyesSum As Variant
            patternStr = "Vat Dyes"
            
            vatDyesSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 23) = vatDyesSum
            
            
            ' Column 24 no import
            
            
            Dim sulphurDyesSum As Variant
            patternStr = "Sulphur Dyes"
            
            sulphurDyesSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 25) = sulphurDyesSum
            
            
            ' Column 26 no import
            
            
            Dim stretchWrappingFilmSum As Variant
            patternStr = "Stretch Wrapping Film"
               
            stretchWrappingFilmSum = Application.Run("utilityFunction.upClause8ReportSumAsclassification", arrUpClause8Range, upClause8OddFiltered, patternStr, 13, 18)
            
            returnArr(1, 27) = stretchWrappingFilmSum
    
    
    
            Dim totalSum As Variant
            
            totalSum = returnArr(1, 1) + returnArr(1, 2) + returnArr(1, 3) + returnArr(1, 4) + returnArr(1, 5) + returnArr(1, 6) + returnArr(1, 7) + returnArr(1, 8) + returnArr(1, 9) + returnArr(1, 10) + returnArr(1, 11) + returnArr(1, 12) + returnArr(1, 13) + returnArr(1, 14) + returnArr(1, 15) + returnArr(1, 16) + returnArr(1, 17) + returnArr(1, 18) + returnArr(1, 19) + returnArr(1, 20) + returnArr(1, 21) + returnArr(1, 22) + returnArr(1, 23) + returnArr(1, 24) + returnArr(1, 25) + returnArr(1, 26) + returnArr(1, 27)
    
    
    
            ' put result to report sheet
    
            Workbooks.Open fileName:=ActiveWorkbook.path & Application.PathSeparator & "Pioneer Denim Limited_Calculation Sheet.xlsx", ReadOnly:=False
            
            Worksheets(2).Select
            
            Dim selectedSheetRange As Range
        
            Set selectedSheetRange = ActiveSheet.Range("e" & ActiveSheet.Cells.Find(upName, LookAt:=xlWhole).Row).Resize(1, 27)
            selectedSheetRange.Activate
            selectedSheetRange = returnArr
            
            
            Dim Result As Variant
            
            Result = Round(Round(totalSum, 2) - Round(Round(upClause8OddFiltered(UBound(upClause8OddFiltered, 1), 18), 2) - Round(yarnUsedSum, 2), 2), 2)
            
            Result = Round(Result)
            
            If Result <> 0 Then
    
                MsgBox "Mismatch " & "UP " & upName & " " & Result
                    
            End If
            
End Function
       
    




Private Function putDyesCatagoryAsImportPerformance(arrUpClause8Range As Range, upName As Variant, sourceDataImportPerformanceDyes As Variant) As Variant
    '    only for put dyes catagory as import performance
        Dim arrUpClause8 As Variant
        
        arrUpClause8 = arrUpClause8Range.value
        
        Dim upClause8OddFiltered As Variant
        
        upClause8OddFiltered = Application.Run("utilityFunction.evenOrOddIndexArrayFilter", arrUpClause8, "odd", True)
        
        
        Dim filterArr As Variant
    
        filterArr = Application.Run("utilityFunction.towDimensionalArrayFilter", upClause8OddFiltered, "Dyes", 13)
        
        Dim mushakOrBillofEntryIndex As Integer
        

        Dim i As Integer

        For i = 1 To UBound(filterArr, 1)

            mushakOrBillofEntryIndex = Application.Run("utilityFunction.indexOf", sourceDataImportPerformanceDyes, filterArr(i, 6), 3, 1, UBound(sourceDataImportPerformanceDyes, 1))

            arrUpClause8Range.Range("M" & filterArr(i, 24) * 2 - 1).value = sourceDataImportPerformanceDyes(mushakOrBillofEntryIndex, 6)

        Next i


        
End Function


    



Private Function putAllCatagoryAsImportPerformance(arrUpClause8Range As Range, upName As Variant, sourceDataImportPerformanceAll As Variant) As Variant
    '    put all catagory as import performance
        Dim regex As New RegExp
        regex.Global = True
        regex.MultiLine = True

        Dim arrUpClause8 As Variant
        
        arrUpClause8 = arrUpClause8Range.value
        
        Dim upClause8OddFiltered, upClause8EvenFiltered As Variant
        
        upClause8OddFiltered = Application.Run("utilityFunction.evenOrOddIndexArrayFilter", arrUpClause8, "odd", True)
        
        upClause8EvenFiltered = Application.Run("utilityFunction.evenOrOddIndexArrayFilter", arrUpClause8, "even", True)

        Dim filterByLc As Variant

        Dim mushakOrBillofEntryIndex, mushakOrBillofEntryIndexFromfilterByLc As Integer

        Dim mushakOrBillOfEntry As Variant

        Dim temp As Variant
        
        
        Dim i As Integer

        For i = 1 To UBound(upClause8OddFiltered, 1) - 1

            
            temp = Trim(upClause8OddFiltered(i, 1))

            regex.pattern = "\d+"

            Set temp = regex.Execute(temp)

            If temp.Count = 0 Then

                filterByLc = Application.Run("utilityFunction.towDimensionalArrayFilter", sourceDataImportPerformanceAll, Trim(upClause8OddFiltered(i, 6)), 3) ' if LC not exist then filter by bill of entry or mushak
            
            Else
            
                temp = temp.Item(0)
    
                filterByLc = Application.Run("utilityFunction.towDimensionalArrayFilter", sourceDataImportPerformanceAll, temp, 4)

            End If

    
            regex.pattern = "\d+$"
            Set mushakOrBillOfEntry = regex.Execute(upClause8OddFiltered(i, 6))
            mushakOrBillOfEntry = mushakOrBillOfEntry.Item(0)

            mushakOrBillofEntryIndexFromfilterByLc = Application.Run("utilityFunction.indexOfMushakOrBillOfEntryWithCompareQtyAndValue", mushakOrBillOfEntry, 3, upClause8OddFiltered(i, 15), 7, upClause8EvenFiltered(i, 15), 8, filterByLc)

            mushakOrBillofEntryIndex = filterByLc(mushakOrBillofEntryIndexFromfilterByLc, UBound(filterByLc, 2))

            arrUpClause8Range.Range("M" & i * 2 - 1).value = sourceDataImportPerformanceAll(mushakOrBillofEntryIndex, 6)

        Next i


        
End Function


Private Function clause8ConvertToCurrentFormat(arr As Variant) As Variant
    ' this function convert clause 8 previous format to current format
    Dim addExtraRow As Variant
    Dim blankArr As Variant
    ReDim blankArr(1 To 2, 1 To UBound(arr, 2))
    
    addExtraRow = Application.Run("utilityFunction.mergeTow2DArrays", arr, blankArr) 'solved for multiple time addSpecificColumnValueOneRowUpTo2DArrayOnLastColumn
    
    Dim afterValueOneRowUp As Variant
    afterValueOneRowUp = Application.Run("utilityFunction.addSpecificColumnValueOneRowUpTo2DArrayOnLastColumn", addExtraRow, 15)
    
    Dim afterPreviousUsedValueOneRowUp As Variant
    afterPreviousUsedValueOneRowUp = Application.Run("utilityFunction.addSpecificColumnValueOneRowUpTo2DArrayOnLastColumn", afterValueOneRowUp, 16)
    
    Dim afterCurrentStockValueOneRowUp As Variant
    afterCurrentStockValueOneRowUp = Application.Run("utilityFunction.addSpecificColumnValueOneRowUpTo2DArrayOnLastColumn", afterPreviousUsedValueOneRowUp, 17)
    
    Dim afterUsedThisUpValueOneRowUp As Variant
    afterUsedThisUpValueOneRowUp = Application.Run("utilityFunction.addSpecificColumnValueOneRowUpTo2DArrayOnLastColumn", afterCurrentStockValueOneRowUp, 18)
    
    Dim afterRemainingValueOneRowUp As Variant
    afterRemainingValueOneRowUp = Application.Run("utilityFunction.addSpecificColumnValueOneRowUpTo2DArrayOnLastColumn", afterUsedThisUpValueOneRowUp, 19)
    
    Dim afterAddExtraColumn As Variant
    afterAddExtraColumn = Application.Run("utilityFunction.AddExtraColumns", afterRemainingValueOneRowUp, 2)
    
    Dim tempSwapping As Variant
    tempSwapping = Application.Run("utilityFunction.SwapColumns", afterAddExtraColumn, 16, 20)
    tempSwapping = Application.Run("utilityFunction.SwapColumns", tempSwapping, 17, 20)
    tempSwapping = Application.Run("utilityFunction.SwapColumns", tempSwapping, 18, 21)
    tempSwapping = Application.Run("utilityFunction.SwapColumns", tempSwapping, 19, 20)
    tempSwapping = Application.Run("utilityFunction.SwapColumns", tempSwapping, 20, 22)
    tempSwapping = Application.Run("utilityFunction.SwapColumns", tempSwapping, 22, 23)
    tempSwapping = Application.Run("utilityFunction.SwapColumns", tempSwapping, 23, 25)
    tempSwapping = Application.Run("utilityFunction.SwapColumns", tempSwapping, 24, 26)
    
    Dim upClause8OddFilteredAfterConvertFormat As Variant
    upClause8OddFilteredAfterConvertFormat = Application.Run("utilityFunction.evenOrOddIndexArrayFilter", tempSwapping, "odd", True)
    
    Dim trimLastRow As Variant
    trimLastRow = Application.Run("utilityFunction.cropedArry", upClause8OddFilteredAfterConvertFormat, 1, UBound(upClause8OddFilteredAfterConvertFormat, 1) - 1)
    
    clause8ConvertToCurrentFormat = trimLastRow
    
End Function
    
  



Private Function putAllUpToHelperFile(arrUpClause8Range As Range, upName As Variant, helperFileWs As Worksheet, isColumn8CurrentFormat As Boolean) As Variant
    'UP clause 8 put to helper file

        Dim arrUpClause8 As Variant
        
        arrUpClause8 = arrUpClause8Range.value
        
        Dim afterConvertedClause8 As Variant
        
        If isColumn8CurrentFormat Then
            afterConvertedClause8 = arrUpClause8
        Else
            afterConvertedClause8 = Application.Run("utilityFunction.clause8ConvertToCurrentFormat", arrUpClause8)
        End If
        
        
        Dim arrUpClause8WithUpNoOnLastColumn As Variant

        arrUpClause8WithUpNoOnLastColumn = Application.Run("utilityFunction.addSpecificStringTo2DArrayOnLastColumn", afterConvertedClause8, "UP-" & upName)



        Dim helperFilePutRange As Range
        Dim helperFilePutRangeLastRow As Variant

        helperFilePutRangeLastRow = helperFileWs.Cells.SpecialCells(xlCellTypeLastCell).Row
        Set helperFilePutRange = helperFileWs.Range("a" & helperFilePutRangeLastRow + 1).Resize(UBound(arrUpClause8WithUpNoOnLastColumn, 1), UBound(arrUpClause8WithUpNoOnLastColumn, 2))
        helperFilePutRange = arrUpClause8WithUpNoOnLastColumn


        
End Function


    Private Function addAllUpToFinalDbDictionary(arrUpClause8Range As Range, upName As Variant, isColumn8CurrentFormat As Boolean, allUpFinalResultDict As Object, totalClassifiedDictKeys As Variant, useGroupDict As Object, impBillAndMushakDb As Object, inputTxt As Variant) As Variant
        'UP clause 8 add to final db dictionary

            Dim arrUpClause8 As Variant

            arrUpClause8 = arrUpClause8Range.value

            Dim afterConvertedClause8 As Variant

            If isColumn8CurrentFormat Then
                afterConvertedClause8 = arrUpClause8
            Else
                afterConvertedClause8 = Application.Run("utilityFunction.clause8ConvertToCurrentFormat", arrUpClause8)
            End If
            
            
            Dim individualUpResultDict As Object
            Set individualUpResultDict = CreateObject("Scripting.Dictionary")
            
            Set individualUpResultDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", individualUpResultDict, 0, totalClassifiedDictKeys)
            
            Dim i As Long
            Dim rawMaterialsAfterRemovedInvalidChr As Variant
            Dim tempClassifiedId As Variant
            Dim dicKeyBillOrMushak As Variant
            Dim qtySum As Variant
            Dim rawMaterialsNotFount As Variant
            Dim descriptionTakeFromUpOrImpPerformance As Variant
            
            qtySum = 0
            rawMaterialsNotFount = ""
            

            descriptionTakeFromUpOrImpPerformance = inputTxt
            
            If descriptionTakeFromUpOrImpPerformance = "UP" Then
                descriptionTakeFromUpOrImpPerformance = True
            Else
                descriptionTakeFromUpOrImpPerformance = False
            End If
            
            For i = LBound(afterConvertedClause8) To UBound(afterConvertedClause8) - 1
            
                        
                If descriptionTakeFromUpOrImpPerformance Then    ' description take from UP or import performance
                
                    If afterConvertedClause8(i, 13) = "Yarn" Or afterConvertedClause8(i, 13) = "Dyes" Or afterConvertedClause8(i, 13) = "" Then
                        ' In previous foramated UP yarn & dyes not defined so, description pick from import performance
                        dicKeyBillOrMushak = Application.Run("general_utility_functions.dictKeyGeneratorWithMushakOrBillOfEntryQtyAndValue", afterConvertedClause8(i, 6), afterConvertedClause8(i, 15), afterConvertedClause8(i, 16))
                        rawMaterialsAfterRemovedInvalidChr = Application.Run("general_utility_functions.RemoveInvalidChars", impBillAndMushakDb(dicKeyBillOrMushak)("Description"))
                
                        If InStr(LCase$(rawMaterialsAfterRemovedInvalidChr), "yarn") Then ' if yarn exist then go inside
                            If Left$(Trim$(afterConvertedClause8(i, 6)), 2) = "C-" Then ' for foreign or local yarn classified
                                rawMaterialsAfterRemovedInvalidChr = Application.Run("general_utility_functions.RemoveInvalidChars", "Foreign Yarn")
                            Else
                                rawMaterialsAfterRemovedInvalidChr = Application.Run("general_utility_functions.RemoveInvalidChars", "Local Yarn")
                            End If
                        End If
                    Else
                        rawMaterialsAfterRemovedInvalidChr = Application.Run("general_utility_functions.RemoveInvalidChars", afterConvertedClause8(i, 13))
                    End If
                
                Else
                
                    If afterConvertedClause8(i, 15) > 0 Then ' if qty. = 0
                
                
                        dicKeyBillOrMushak = Application.Run("general_utility_functions.dictKeyGeneratorWithMushakOrBillOfEntryQtyAndValue", afterConvertedClause8(i, 6), afterConvertedClause8(i, 15), afterConvertedClause8(i, 16))
                        rawMaterialsAfterRemovedInvalidChr = Application.Run("general_utility_functions.RemoveInvalidChars", impBillAndMushakDb(dicKeyBillOrMushak)("Description"))
                    
                        If InStr(LCase$(rawMaterialsAfterRemovedInvalidChr), "yarn") Then ' if yarn exist then go inside
                            If Left$(Trim$(afterConvertedClause8(i, 6)), 2) = "C-" Then ' for foreign or local yarn classified
                                rawMaterialsAfterRemovedInvalidChr = Application.Run("general_utility_functions.RemoveInvalidChars", "Foreign Yarn")
                            Else
                                rawMaterialsAfterRemovedInvalidChr = Application.Run("general_utility_functions.RemoveInvalidChars", "Local Yarn")
                            End If
                        End If
                        
                    
                    Else
                        rawMaterialsAfterRemovedInvalidChr = Application.Run("general_utility_functions.RemoveInvalidChars", afterConvertedClause8(i, 13))
                    End If
                
                End If
            
                                
                tempClassifiedId = useGroupDict(rawMaterialsAfterRemovedInvalidChr)
                
                If individualUpResultDict.Exists(tempClassifiedId) Then
                    qtySum = qtySum + afterConvertedClause8(i, 21) ' take sum for compare with total sum
                    individualUpResultDict(tempClassifiedId) = individualUpResultDict(tempClassifiedId) + afterConvertedClause8(i, 21)
                Else
                    'not found raw materials here
                    If descriptionTakeFromUpOrImpPerformance Then
                    
                        allUpFinalResultDict("rawMaterialsNotClassifiedDict")(rawMaterialsAfterRemovedInvalidChr) = afterConvertedClause8(i, 13)
                        
                        rawMaterialsNotFount = rawMaterialsNotFount & afterConvertedClause8(i, 13) & ", "
                    Else
                        If afterConvertedClause8(i, 15) > 0 Then ' if qty. = 0
                            allUpFinalResultDict("rawMaterialsNotClassifiedDict")(rawMaterialsAfterRemovedInvalidChr) = impBillAndMushakDb(dicKeyBillOrMushak)("Description")
                            
                            rawMaterialsNotFount = rawMaterialsNotFount & impBillAndMushakDb(dicKeyBillOrMushak)("Description") & ", "
                        Else
                            allUpFinalResultDict("rawMaterialsNotClassifiedDict")(rawMaterialsAfterRemovedInvalidChr) = afterConvertedClause8(i, 13)
                            
                            rawMaterialsNotFount = rawMaterialsNotFount & afterConvertedClause8(i, 13) & ", "
                        End If
                    End If

                End If
                
            Next i
            
            
            If Not Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", Round(qtySum, 2), Round(afterConvertedClause8(UBound(afterConvertedClause8), 21), 2), 0.5) Then
                MsgBox "In UP " & upName & " total sum mismatch = " & Round(Round(qtySum, 2) - Round(afterConvertedClause8(UBound(afterConvertedClause8), 21), 2), 2) & Chr(10) & "Materials not found " & rawMaterialsNotFount
            End If
            
            allUpFinalResultDict.Add upName, individualUpResultDict
            
            Set addAllUpToFinalDbDictionary = allUpFinalResultDict
            
    End Function




    
Private Function AddExtraColumns(arr As Variant, numExtraColumns As Integer) As Variant
' AddExtraColumns takes the 2-D array arr and the numExtraColumns integer as inputs.
' It creates a new array resultArr with the desired number of extra columns,
' and copies the values from the original array into the new array.
    Dim numRows As Long
    Dim numCols As Long
    Dim resultArr As Variant
    Dim i As Long, j As Long
    
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    ReDim resultArr(1 To numRows, 1 To numCols + numExtraColumns)
    
    For i = 1 To numRows
        For j = 1 To numCols
            resultArr(i, j) = arr(i, j)
        Next j
    Next i
    
    AddExtraColumns = resultArr
End Function


Private Function SwapColumns(arr As Variant, column1 As Integer, column2 As Integer) As Variant
' this function takes a 2-D array and two integer values representing the column numbers to swap.
' The function will swap the values in the specified columns and return the new array
    Dim numRows As Long
    Dim numCols As Long
    Dim resultArr As Variant
    Dim i As Long, j As Long
    
    numRows = UBound(arr, 1)
    numCols = UBound(arr, 2)
    
    ReDim resultArr(1 To numRows, 1 To numCols)
    
    For i = 1 To numRows
        For j = 1 To numCols
            resultArr(i, j) = arr(i, j)
        Next j
        
        ' Swap column values
        Dim temp As Variant
        temp = resultArr(i, column1)
        resultArr(i, column1) = resultArr(i, column2)
        resultArr(i, column2) = temp
    Next i
    
    SwapColumns = resultArr
End Function


Private Function DoesStringExistInWorksheets(searchString As String, searchStringWs As Worksheet) As Boolean
'    this function return true if provided string find in provided worksheet else return false
    Dim ws As Worksheet
    Dim rng As Range
    Dim found As Boolean
    
    found = False ' Initialize the flag
    
    Set ws = searchStringWs
    Set rng = ws.UsedRange
    
    If Not rng.Find(What:=searchString, LookIn:=xlValues, LookAt:=xlPart) Is Nothing Then
        found = True
    End If

    DoesStringExistInWorksheets = found
End Function


Private Function isCompareValuesLessThanProvidedValue(num1 As Variant, num2 As Variant, differenceCompare As Variant) As Boolean
'    "num1", "num2" are the compare value and "differenceCompare" is the value what we want to know "is less than this value"
    Dim difference As Variant
    difference = Abs(num1 - num2)
    
    If difference < differenceCompare Then
        isCompareValuesLessThanProvidedValue = True
    Else
        isCompareValuesLessThanProvidedValue = False
    End If
End Function


Private Function putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFile(workingWs As Worksheet, billOfEntryOrMushakColumn As Integer, qtyColumn As Integer, valueColumn As Integer, usedQtyColumn As Integer, usedValueColumn As Integer, remarkColumn As Integer, afterMergedClause8OfAllUp As Variant) As Variant
    'Used Qty & Value put to import performance file

    Dim regex As New RegExp
    regex.Global = True
    regex.MultiLine = True
    regex.pattern = ".+"
    Dim mushakOrBillOfEntry As Variant

    Dim workingRange As Range

    workingWs.AutoFilterMode = False

    Set workingRange = workingWs.Range("A6:" & "AB" & workingWs.Range("C6").End(xlDown).Row)

    Dim workingRangeValueArr As Variant

    workingRangeValueArr = workingRange.value

    Dim totalUsedQtyArr, totalUsedValueArr, remarkArr As Variant

    ReDim totalUsedQtyArr(1 To UBound(workingRangeValueArr, 1), 1 To 1)

    ReDim totalUsedValueArr(1 To UBound(workingRangeValueArr, 1), 1 To 1)

    ReDim remarkArr(1 To UBound(workingRangeValueArr, 1), 1 To 1)

    Dim tempQty, tempValue, tempRemark As Variant

    Dim filterByBillOfEntryOrMushak, qtyFromImportPerformance, valueFromImportPerformance As Variant

    Dim i As Long

    For i = 1 To UBound(workingRangeValueArr, 1)

        qtyFromImportPerformance = workingRangeValueArr(i, qtyColumn)
        valueFromImportPerformance = workingRangeValueArr(i, valueColumn)

        Set mushakOrBillOfEntry = regex.Execute(workingRangeValueArr(i, billOfEntryOrMushakColumn))
        mushakOrBillOfEntry = mushakOrBillOfEntry.Item(0)

        filterByBillOfEntryOrMushak = Application.Run("utilityFunction.filterMushakOrBillOfEntryArrayWithCompareQtyAndValue", mushakOrBillOfEntry, 6, qtyFromImportPerformance, 15, valueFromImportPerformance, 16, afterMergedClause8OfAllUp)

        If IsArray(filterByBillOfEntryOrMushak) Then

            tempQty = Application.Run("utilityFunction.sumArrColumn", filterByBillOfEntryOrMushak, 21)
            tempValue = Application.Run("utilityFunction.sumArrColumn", filterByBillOfEntryOrMushak, 22)
            tempRemark = Application.Run("utilityFunction.concatSpecificColumnString", filterByBillOfEntryOrMushak, UBound(filterByBillOfEntryOrMushak, 2) - 1)

            If Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", qtyFromImportPerformance, tempQty, 0.8) Then

                tempQty = qtyFromImportPerformance

            End If

            If Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", valueFromImportPerformance, tempValue, 0.8) Then

                tempValue = valueFromImportPerformance

            End If

        Else

            tempQty = Null
            tempValue = Null
            tempRemark = Null

        End If

        totalUsedQtyArr(i, 1) = tempQty
        totalUsedValueArr(i, 1) = tempValue
        remarkArr(i, 1) = tempRemark

    Next i

    workingRange.Columns(usedQtyColumn) = totalUsedQtyArr
    workingRange.Columns(usedValueColumn) = totalUsedValueArr
    workingRange.Columns(remarkColumn) = remarkArr

End Function

Private Function putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFileWithJson(workingWs As Worksheet, billOfEntryOrMushakColumn As Integer, lcColumn As Integer, qtyColumn As Integer, valueColumn As Integer, usedQtyColumn As Integer, usedValueColumn As Integer, remarkColumn As Integer, allUpClause8UseAsMushakOrBillOfEntryDic As Variant) As Variant
    'Used Qty & Value put to import performance file

    Dim workingRange As Range

    workingWs.AutoFilterMode = False

    If IsEmpty(workingWs.Range("C7").Value) Then

        Set workingRange = workingWs.Range("A6:AB6")

    Else

        Set workingRange = workingWs.Range("A6:" & "AB" & workingWs.Range("C6").End(xlDown).Row)

    End If

    Dim workingRangeValueArr As Variant

    workingRangeValueArr = workingRange.value

    Dim totalUsedQtyArr, totalUsedValueArr, remarkArr As Variant

    ReDim totalUsedQtyArr(1 To UBound(workingRangeValueArr, 1), 1 To 1)

    ReDim totalUsedValueArr(1 To UBound(workingRangeValueArr, 1), 1 To 1)

    ReDim remarkArr(1 To UBound(workingRangeValueArr, 1), 1 To 1)

    Dim tempQty, tempValue, tempRemark As Variant

    Dim qtyFromImportPerformance, valueFromImportPerformance As Variant

    Dim tempMuOrBillKey As String

    Dim i As Long

    For i = 1 To UBound(workingRangeValueArr, 1)

        qtyFromImportPerformance = workingRangeValueArr(i, qtyColumn)
        valueFromImportPerformance = workingRangeValueArr(i, valueColumn)

        tempMuOrBillKey = Application.Run("general_utility_functions.dictKeyGeneratorWithLcMushakOrBillOfEntryQtyAndValue", workingRangeValueArr(i, lcColumn), workingRangeValueArr(i, billOfEntryOrMushakColumn), workingRangeValueArr(i, qtyColumn), workingRangeValueArr(i, valueColumn))

        If allUpClause8UseAsMushakOrBillOfEntryDic.Exists(tempMuOrBillKey) Then

            tempQty = allUpClause8UseAsMushakOrBillOfEntryDic(tempMuOrBillKey)("sumOfAllUpUsedQty")
            tempValue = allUpClause8UseAsMushakOrBillOfEntryDic(tempMuOrBillKey)("sumOfAllUpUsedValue")
            tempRemark = Replace(Replace(allUpClause8UseAsMushakOrBillOfEntryDic(tempMuOrBillKey)("usedUpList"), ",", "", 1, 1), ",", Chr(10))

            If Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", qtyFromImportPerformance, tempQty, 0.8) Then

                tempQty = qtyFromImportPerformance

            End If

            If Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", valueFromImportPerformance, tempValue, 0.8) Then

                tempValue = valueFromImportPerformance

            End If

        Else

            tempQty = Null
            tempValue = Null
            tempRemark = Null

        End If

        totalUsedQtyArr(i, 1) = tempQty
        totalUsedValueArr(i, 1) = tempValue
        remarkArr(i, 1) = tempRemark

    Next i

    workingRange.Columns(usedQtyColumn) = totalUsedQtyArr
    workingRange.Columns(usedValueColumn) = totalUsedValueArr
    workingRange.Columns(remarkColumn) = remarkArr

End Function


Private Function CombinedAllSheetsMushakOrBillOfEntryDbDict(importPerformanceFilePath As String) As Object
' this function combined all import performance sheets for short cut one call. if import performance criteria or column changed this function should modified
    Application.ScreenUpdating = False
    Dim importPerformanceWb As Workbook
    Set importPerformanceWb = Workbooks.Open(importPerformanceFilePath)
    
    Dim yarnImportWs As Worksheet
    Set yarnImportWs = importPerformanceWb.Worksheets("Yarn (Import)")
    
    Dim yarnLocalWs As Worksheet
    Set yarnLocalWs = importPerformanceWb.Worksheets("Yarn (Local)")
    
    Dim dyesWs As Worksheet
    Set dyesWs = importPerformanceWb.Worksheets("Dyes")
    
    Dim chemicalsImportWs As Worksheet
    Set chemicalsImportWs = importPerformanceWb.Worksheets("Chemicals (Import)")
    
    Dim chemicalsLocalWs As Worksheet
    Set chemicalsLocalWs = importPerformanceWb.Worksheets("Chemicals (Local)")
    
    Dim stretchWrappingFilmWs As Worksheet
    Set stretchWrappingFilmWs = importPerformanceWb.Worksheets("St.Wrap.Film (Import)")
    
    Dim sourceDataImportPerformanceYarnImport As Variant

    If IsEmpty(yarnImportWs.Range("C7").Value) Then

        sourceDataImportPerformanceYarnImport = yarnImportWs.Range("A6:N6").value

    Else

        sourceDataImportPerformanceYarnImport = yarnImportWs.Range("A6:" & "N" & yarnImportWs.Range("C6").End(xlDown).Row).value

    End If
    
    Dim sourceDataImportPerformanceYarnLocal As Variant

    If IsEmpty(yarnLocalWs.Range("C7").Value) Then

        sourceDataImportPerformanceYarnLocal = yarnLocalWs.Range("A6:N6").value

    Else

        sourceDataImportPerformanceYarnLocal = yarnLocalWs.Range("A6:" & "N" & yarnLocalWs.Range("C6").End(xlDown).Row).value

    End If
    
    Dim sourceDataImportPerformanceDyes As Variant

    If IsEmpty(dyesWs.Range("C7").Value) Then

        sourceDataImportPerformanceDyes = dyesWs.Range("A6:N6").value

    Else

        sourceDataImportPerformanceDyes = dyesWs.Range("A6:" & "N" & dyesWs.Range("C6").End(xlDown).Row).value

    End If
    
    Dim sourceDataImportPerformanceChemicalsImport As Variant

    If IsEmpty(chemicalsImportWs.Range("C7").Value) Then

        sourceDataImportPerformanceChemicalsImport = chemicalsImportWs.Range("A6:N6").value

    Else

        sourceDataImportPerformanceChemicalsImport = chemicalsImportWs.Range("A6:" & "N" & chemicalsImportWs.Range("C6").End(xlDown).Row).value

    End If
    
    Dim sourceDataImportPerformanceChemicalsLocal As Variant

    If IsEmpty(chemicalsLocalWs.Range("C7").Value) Then

        sourceDataImportPerformanceChemicalsLocal = chemicalsLocalWs.Range("A6:N6").value

    Else

        sourceDataImportPerformanceChemicalsLocal = chemicalsLocalWs.Range("A6:" & "N" & chemicalsLocalWs.Range("C6").End(xlDown).Row).value

    End If
    
    Dim sourceDataImportPerformanceStretchWrappingFilm As Variant
    
    If IsEmpty(stretchWrappingFilmWs.Range("C7").Value) Then

        sourceDataImportPerformanceStretchWrappingFilm = stretchWrappingFilmWs.Range("A6:N6").value

    Else

        sourceDataImportPerformanceStretchWrappingFilm = stretchWrappingFilmWs.Range("A6:" & "N" & stretchWrappingFilmWs.Range("C6").End(xlDown).Row).value

    End If

    importPerformanceWb.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
    Dim importPerformanceDbDict As Object
    Set importPerformanceDbDict = CreateObject("Scripting.Dictionary")
    
    Set importPerformanceDbDict = Application.Run("dictionary_utility_functions.CreateMushakOrBillOfEntryDbDict", importPerformanceDbDict, sourceDataImportPerformanceYarnImport, 4, 3, 7, 8, 6, Array("BillOfEntryOrMushak", "LC", "HSCode", "Description", "Qty", "Value", "UsedQty", "UsedValue", "BalanceQty", "BalanceValue"), Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12))
    Set importPerformanceDbDict = Application.Run("dictionary_utility_functions.CreateMushakOrBillOfEntryDbDict", importPerformanceDbDict, sourceDataImportPerformanceYarnLocal, 4, 3, 7, 8, 6, Array("BillOfEntryOrMushak", "LC", "HSCode", "Description", "Qty", "Value", "UsedQty", "UsedValue", "BalanceQty", "BalanceValue"), Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12))
    Set importPerformanceDbDict = Application.Run("dictionary_utility_functions.CreateMushakOrBillOfEntryDbDict", importPerformanceDbDict, sourceDataImportPerformanceDyes, 4, 3, 7, 8, 6, Array("BillOfEntryOrMushak", "LC", "HSCode", "Description", "Qty", "Value", "UsedQty", "UsedValue", "BalanceQty", "BalanceValue"), Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12))
    Set importPerformanceDbDict = Application.Run("dictionary_utility_functions.CreateMushakOrBillOfEntryDbDict", importPerformanceDbDict, sourceDataImportPerformanceChemicalsImport, 4, 3, 8, 9, 7, Array("BillOfEntryOrMushak", "LC", "HSCode", "Description", "Qty", "Value", "UsedQty", "UsedValue", "BalanceQty", "BalanceValue"), Array(3, 4, 5, 7, 8, 9, 10, 11, 12, 13))
    Set importPerformanceDbDict = Application.Run("dictionary_utility_functions.CreateMushakOrBillOfEntryDbDict", importPerformanceDbDict, sourceDataImportPerformanceChemicalsLocal, 4, 3, 8, 9, 7, Array("BillOfEntryOrMushak", "LC", "HSCode", "Description", "Qty", "Value", "UsedQty", "UsedValue", "BalanceQty", "BalanceValue"), Array(3, 4, 5, 7, 8, 9, 10, 11, 12, 13))
    Set importPerformanceDbDict = Application.Run("dictionary_utility_functions.CreateMushakOrBillOfEntryDbDict", importPerformanceDbDict, sourceDataImportPerformanceStretchWrappingFilm, 4, 3, 8, 9, 7, Array("BillOfEntryOrMushak", "LC", "HSCode", "Description", "Qty", "Value", "UsedQty", "UsedValue", "BalanceQty", "BalanceValue"), Array(3, 4, 5, 7, 8, 9, 10, 11, 12, 13))
    
    Set CombinedAllSheetsMushakOrBillOfEntryDbDict = importPerformanceDbDict
    
End Function

Private Function importPerformanceCommentedBillOfEntryOrMushakDbFromProvidedSheet(ws As worksheet, lcCol As Integer, mushakOrBillOfEntryCol As Integer, qtyCol As Integer, valueCol As Integer) As Object
    'returned all commented bill of entry or mushak dictionary

    ws.AutoFilterMode = False

    Dim workingRange As Range

    If IsEmpty(ws.Range("C7").Value) Then

        Set workingRange = ws.Range("A6:N6")

    Else

        Set workingRange = ws.Range("A6:" & "N" & ws.Range("C6").End(xlDown).Row)

    End If

    Dim commentedBillOfEntryOrMushak As Object
    Set commentedBillOfEntryOrMushak = CreateObject("Scripting.Dictionary")

    Dim tempMuOrBillKey As String

    Dim i As Long

    For i = 1 To workingRange.Rows.Count

        If Not workingRange(i, mushakOrBillOfEntryCol).Comment Is Nothing Then   'check if the cell has a comment

            tempMuOrBillKey = Application.Run("general_utility_functions.dictKeyGeneratorWithLcMushakOrBillOfEntryQtyAndValue", workingRange(i, lcCol), workingRange(i, mushakOrBillOfEntryCol), workingRange(i, qtyCol), workingRange(i, valueCol))

            If commentedBillOfEntryOrMushak.Exists(tempMuOrBillKey) Then

                commentedBillOfEntryOrMushak(tempMuOrBillKey)("Entry_Count") = commentedBillOfEntryOrMushak(tempMuOrBillKey)("Entry_Count") + 1
                commentedBillOfEntryOrMushak(tempMuOrBillKey)("comment") = commentedBillOfEntryOrMushak(tempMuOrBillKey)("comment") & Chr(10) & workingRange(i, mushakOrBillOfEntryCol).Comment.Text

            Else

                commentedBillOfEntryOrMushak.Add tempMuOrBillKey, CreateObject("Scripting.Dictionary")
                commentedBillOfEntryOrMushak(tempMuOrBillKey)("Entry_Count") = 1
                commentedBillOfEntryOrMushak(tempMuOrBillKey)("comment") = workingRange(i, mushakOrBillOfEntryCol).Comment.Text

            End If

        End If

    Next i

    Set importPerformanceCommentedBillOfEntryOrMushakDbFromProvidedSheet = commentedBillOfEntryOrMushak

End Function

Private Function upSequenceStrGenerator(upArr As Variant, sequenceInnerText As String, sequenceBreakCharacterCode As Long) As String
    'this function received an UP No. array and return UP sequence string
    Dim uPSequenceObj As Object
    Dim uPSequenceStr As String
    Dim tempStr As String

    uPSequenceStr = ""

    Set uPSequenceObj = Application.Run("Sorting_Algorithms.SplituPSequence", upArr)

    Dim dictKey As Variant

    For Each dictKey In uPSequenceObj.keys

        If uPSequenceObj(dictKey)("sequenceStart") = uPSequenceObj(dictKey)("sequenceEnd") Then

            tempStr = uPSequenceObj(dictKey)("sequenceStart")

        Else
                
            tempStr = uPSequenceObj(dictKey)("sequenceStart") & sequenceInnerText & uPSequenceObj(dictKey)("sequenceEnd")
            
        End If

        uPSequenceStr = uPSequenceStr & tempStr & Chr(sequenceBreakCharacterCode)

    Next dictKey

    upSequenceStrGenerator = Left(uPSequenceStr, Len(uPSequenceStr) - 1)
    
End Function

Private Function cellsMarkingAsValue(markingRange As Range, criteriaValue As String)
    
    Dim eachCell As Range

    For Each eachCell In markingRange

        If eachCell.Value = criteriaValue Then

            eachCell.Interior.Color = RGB(255, 0, 0)

        End If

    Next eachCell
    
End Function