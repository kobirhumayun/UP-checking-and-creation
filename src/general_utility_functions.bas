Attribute VB_Name = "general_utility_functions"
Option Explicit


Private Function InsertStringAtPosition(originalString As String, insertString As String, position As Integer) As Variant
    Dim length As Integer
    length = Len(originalString)
    
    If length >= position Then
        InsertStringAtPosition = Left(originalString, length - (position - 1)) & insertString & Right(originalString, (position - 1))
    Else
        ' Handle the case where the original string is shorter than 5 characters
        InsertStringAtPosition = Null
    End If
End Function


Private Function RemoveInvalidChars(ByVal inputString As String) As String
    Dim invalidChars As String
    invalidChars = " ~`!@#$%^&*()-+=[]\{}|;':"",./<>?" & vbNewLine & Chr(10) & Chr(160)
    
    Dim resultString As String
    Dim i As Long
    
    For i = 1 To Len(inputString)
        Dim currentChar As String
        currentChar = Mid(inputString, i, 1)
        
        If InStr(invalidChars, currentChar) = 0 Then
            resultString = resultString & currentChar
        End If
    Next i
    
    RemoveInvalidChars = resultString
End Function


Private Function oneDArrayConvertToTwoDArray(inputArray As Variant) As Variant
    Dim outputArray As Variant

    ReDim outputArray(LBound(inputArray) To UBound(inputArray), 1 To 1)

    Dim i As Long
    For i = LBound(inputArray) To UBound(inputArray)
        outputArray(i, 1) = inputArray(i)
    Next i

    oneDArrayConvertToTwoDArray = outputArray
End Function

Private Function ExtractStringLeftOfComma(ByVal inputText As String) As Variant

    Dim commaPosition As Long
    Dim extractedString As String

    ' Find the position of the first comma
    commaPosition = InStr(inputText, ",")

    If commaPosition > 0 Then
        ' Extract the string left of the first comma
        extractedString = Left(inputText, commaPosition - 1)
        ExtractStringLeftOfComma = extractedString
    Else
        ExtractStringLeftOfComma = Null
    End If

End Function


Private Function regExReturnedObj(str As Variant, pattern As Variant, isGlobal As Boolean, isIgnoreCase As Boolean, isMultiLine As Boolean) As Object

    Dim regex As Object

    ' Convert the str to a string
    str = CStr(str)

    ' Convert the pattern to a string
    pattern = CStr(pattern)

    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .MultiLine = isMultiLine
        .pattern = pattern
    End With

    ' Return the test result
    Set regExReturnedObj = regex.Execute(str)

End Function


Private Function isStrPatternExist(str As Variant, pattern As Variant, isGlobal As Boolean, isIgnoreCase As Boolean, isMultiLine As Boolean) As Boolean

    Dim regex As Object

    ' Convert the str to a string
    str = CStr(str)

    ' Convert the pattern to a string
    pattern = CStr(pattern)

    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        .MultiLine = isMultiLine
        .pattern = pattern
    End With

    ' Return the test result
    isStrPatternExist = regex.test(str)

End Function

Private Function extractAndFormatUdNo(udNo As String) As String
    'this function extract and formated UD, Sample formt "UDNo_Year" or "UDNo_Year_AM"

    Dim tempObj As Object
    Dim uDyear As String
    Dim formatedUd As String
    Dim tempUd As String

    Set tempObj = Application.Run("general_utility_functions.regExReturnedObj", udNo, "(BGMEA\/DHK\/UD\/\d+)|(BGMEA\/DHK\/AM\/\d+)", True, True, True) ' for UD and Amnd Year

    uDyear = Right$(tempObj(0), 4)

    Set tempObj = Application.Run("general_utility_functions.regExReturnedObj", udNo, "(\d+$)|(\d+\-\d+$)", True, True, True) ' for UD and Amnd No.

    tempUd = tempObj(0)

    Set tempObj = Application.Run("general_utility_functions.regExReturnedObj", tempUd, "\d+", True, True, True) ' extract UD and Amnd No.

    If tempObj.Count = 1 Then

        formatedUd = Val(tempObj(0)) & "_" & uDyear

    ElseIf tempObj.Count = 2 Then

        formatedUd = Val(tempObj(0)) & "_" & Val(tempObj(1)) & "_" & uDyear

    End If

    extractAndFormatUdNo = formatedUd

End Function


Private Function ExtractLeftDigitWithRegex(number As Variant) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String
    Dim leftDigit As Variant
    
    ' Convert the number to a string
    Dim numberString As String
    numberString = CStr(number)
    
    ' Define the regular expression pattern to match the left digits
    pattern = "\d+"
    
    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        .pattern = pattern
    End With
    
    ' Get the matches
    Set matches = regex.Execute(numberString)
    
    ' Check if there's a match
    If matches.Count > 0 Then
        
        leftDigit = matches(0)
        
    Else
        ' Default to 0 if no match found
        leftDigit = 0
    End If
    
    ' Return the extracted left digit
    ExtractLeftDigitWithRegex = leftDigit
    
End Function

Private Function ExtractFirstLineWithRegex(str As Variant) As Variant

    Dim matches As Object
    Dim firstLine As Variant
    
    ' Get the matches
    Set matches = Application.Run("general_utility_functions.regExReturnedObj", str, ".+", True, True, True) ' extract first line
    
    ' Check if there's a match
    If matches.Count > 0 Then
        
        firstLine = matches(0)
        
    Else
        ' Default to 0 if no match found
        firstLine = 0
    End If
    
    ' Return the extracted first line
    ExtractFirstLineWithRegex = firstLine
    
End Function

Private Function ExtractRightDigitFromEnd(str As Variant) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String
    Dim rightDigit As Variant

    ' Convert the str to a string
    Dim numberString As String
    numberString = CStr(str)

    ' Define the regular expression pattern to match the right digits
    pattern = "\d+$"

    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .pattern = pattern
    End With

    ' Get the matches
    Set matches = regex.Execute(numberString)

    ' Check if there's a match
    If matches.Count > 0 Then

        rightDigit = matches(0)

    Else
        ' Default to 0 if no match found
        rightDigit = 0
    End If

    ' Return the extracted right digit
    ExtractRightDigitFromEnd = rightDigit

End Function


Private Function ExtractRightDigitOfMuOrBillWithRegex(mushakOrBillOfEntry As Variant) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim pattern As String
    Dim rightDigit As Variant

    ' Convert the mushakOrBillOfEntry to a string
    Dim numberString As String
    numberString = CStr(mushakOrBillOfEntry)

    ' Define the regular expression pattern to match the right digits
    pattern = "\d+$"

    ' Create a RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .pattern = pattern
    End With

    ' Get the matches
    Set matches = regex.Execute(numberString)

    ' Check if there's a match
    If matches.Count > 0 Then

        rightDigit = matches(0)

    Else
        ' Default to 0 if no match found
        rightDigit = 0
    End If

    ' Return the extracted right digit
    ExtractRightDigitOfMuOrBillWithRegex = rightDigit

End Function


Private Function dictKeyGeneratorWithMushakOrBillOfEntryQtyAndValue(mushakOrBillOfEntry As Variant, qty As Variant, value As Variant) As Variant

    
    mushakOrBillOfEntry = Application.Run("general_utility_functions.ExtractRightDigitOfMuOrBillWithRegex", mushakOrBillOfEntry)  'take right digits only for use dic keys
    
    qty = Application.Run("general_utility_functions.ExtractLeftDigitWithRegex", qty)
    
    value = Application.Run("general_utility_functions.ExtractLeftDigitWithRegex", value)

    dictKeyGeneratorWithMushakOrBillOfEntryQtyAndValue = mushakOrBillOfEntry & "_" & qty & "_" & value
    
End Function

Private Function dictKeyGeneratorWithLcMushakOrBillOfEntryQtyAndValue(lc As Variant, mushakOrBillOfEntry As Variant, qty As Variant, value As Variant) As Variant
    
    lc = Application.Run("general_utility_functions.ExtractFirstLineWithRegex", lc)  'take first line only for use dic keys
    
    mushakOrBillOfEntry = Application.Run("general_utility_functions.ExtractRightDigitOfMuOrBillWithRegex", mushakOrBillOfEntry)  'take right digits only for use dic keys
    
    qty = Application.Run("general_utility_functions.ExtractLeftDigitWithRegex", qty)
    
    value = Application.Run("general_utility_functions.ExtractLeftDigitWithRegex", value)

    dictKeyGeneratorWithLcMushakOrBillOfEntryQtyAndValue = lc & "_" & mushakOrBillOfEntry & "_" & qty & "_" & value
    
End Function

Private Function dictKeyGeneratorWithProvidedArrayElements(ByVal arr As Variant) As String

    Dim tempDictKeyStr As String
    Dim elements As Variant

    tempDictKeyStr = ""

    For Each elements In arr
       tempDictKeyStr = tempDictKeyStr & "_" & elements
    Next elements

    tempDictKeyStr = Right(tempDictKeyStr, Len(tempDictKeyStr) - 1)

    dictKeyGeneratorWithProvidedArrayElements = tempDictKeyStr
    
End Function

Private Function upNoAndYearExtrac(upNo As Variant) As Variant
  'this function extract up and year of up
      
      Dim upOnlyNo, upYear As Variant
          
      Dim regex As New RegExp
      regex.Global = True
      regex.MultiLine = True
  
      regex.pattern = "\d+"
      Set upOnlyNo = regex.Execute(upNo)
      upOnlyNo = upOnlyNo.Item(0)
      
      regex.pattern = "\d+$"
      Set upYear = regex.Execute(upNo)
      upYear = upYear.Item(0)
      
      Dim temp As Variant
      ReDim temp(1 To 2)
      temp(1) = upOnlyNo
      temp(2) = upYear
      
      upNoAndYearExtrac = temp
  
End Function

Private Function upNoAndYearExtracAsDict(upNo As Variant) As Variant
    'this function extract up and year of up
      
    Dim onlyUpNo, onlyUpYear As Variant

    Set onlyUpNo = Application.Run("general_utility_functions.regExReturnedObj", upNo, "\d+", True, True, True)

    Set onlyUpYear = Application.Run("general_utility_functions.regExReturnedObj", upNo, "\d+$", True, True, True)

    Dim tempDict As Object
    Set tempDict = CreateObject("Scripting.Dictionary")
    
    tempDict("only_up_no") = onlyUpNo.Item(0)
    tempDict("only_up_year") = onlyUpYear.Item(0)

    Set upNoAndYearExtracAsDict = tempDict
  
End Function
  
Private Function returnSelectedFilesFullPathArr(ByVal initialPath As String) As Variant
  Dim fileDialog As Object
  Dim selectedFiles As Variant
  Dim i As Long
  Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
  With fileDialog
      .Title = "Select Files"
      .AllowMultiSelect = True
       .InitialFileName = initialPath
      If .Show = -1 Then
          ReDim selectedFiles(1 To .SelectedItems.Count)
          For i = 1 To .SelectedItems.Count
              selectedFiles(i) = .SelectedItems.Item(i)
          Next i
      End If
  End With

  returnSelectedFilesFullPathArr = selectedFiles
End Function

Private Function CopyFileToFolderUsingFSO(sourceFilePath As String, targetFolderPath As String, overwrite As Boolean)

    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(sourceFilePath) Then
        Dim fileName As String
        fileName = fso.GetFileName(sourceFilePath)

        Dim targetPath As String
        targetPath = fso.BuildPath(targetFolderPath, fileName)

        fso.CopyFile sourceFilePath, targetPath, overwrite

        ' Check if the copy was successful
        If Err.number = 0 Then
            MsgBox "File " & sourceFilePath & " copied successfully!"
        Else
            MsgBox "Target " & targetFolderPath & " " & Err.Description
        End If
    Else
        MsgBox "Source file " & sourceFilePath & " not found."
    End If

End Function

Private Function CopyFileAsNewFileFSO(sourceFilePath As String, newFilePath As String, overwrite As Boolean)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(sourceFilePath) Then
        
        fso.CopyFile sourceFilePath, newFilePath, overwrite

    Else
        MsgBox "Source file " & sourceFilePath & " not found."
    End If

End Function


Private Function sequentiallyRelateTwoArraysAsDictionary(properties_1 As String, properties_2 As String, properties_1_Arr As Variant, properties_2_Arr As Variant) As Variant
    'this function take two str as properties & two arr contain values then return a dictionary, dictionary use first arr elements as keys
    'and all keys are also dictionaries that contain same sequential elements of two arr

    Dim mainDictionary As Object
    Set mainDictionary = CreateObject("Scripting.Dictionary")

    Dim subDictionary As Object
    Dim removedAllInvalidChrFromKeys As String
    
    If LBound(properties_1_Arr) <> LBound(properties_2_Arr) Or UBound(properties_1_Arr) <> UBound(properties_2_Arr) Then
    
        MsgBox "Both array length are not same"
        Exit Function
        
    End If
    
    Dim i As Long

    ' create sub dictionary and add to main dictionary
    For i = LBound(properties_1_Arr) To UBound(properties_1_Arr)

        Set subDictionary = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", Array(properties_1, properties_2), Array(properties_1_Arr(i), properties_2_Arr(i)))

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", properties_1_Arr(i))   'remove all invalid characters for use dic keys
        mainDictionary.Add removedAllInvalidChrFromKeys, subDictionary
    Next i

    Set sequentiallyRelateTwoArraysAsDictionary = mainDictionary

End Function

Private Function upClause8InformationFromProvidedWs(ws As Worksheet) As Object
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


    Set upClause8InformationFromProvidedWs = upClause8Dic

End Function


Private Function sumUsedQtyAndValueAsMushakOrBillOfEntryFromSelectedUpFile() As Object
    ' this function give sum of used Qty. & value as mushak or bill of entry
    ' data as dictionary from selected previous calculated json text file & selected UP file.
    ' also merged dictionary save as json text file.

    Application.ScreenUpdating = False

    Dim allUpClause8UseAsMushakOrBillOfEntryDic As Object
    Set allUpClause8UseAsMushakOrBillOfEntryDic = CreateObject("Scripting.Dictionary")

    Dim curentUpClause8Dict As Object

    Dim jsonPath As String
    jsonPath = ActiveWorkbook.path & Application.PathSeparator & "json-used-up-clause8"

    Dim initialUpPath As String
    initialUpPath = ActiveWorkbook.path & Application.PathSeparator & "UP-period-2024-2025"

    Dim upPathArr As Variant
    Dim jsonPathArr As Variant

    Dim currentUpWb As Workbook
    Dim currentUpWs As Worksheet

    Dim curentUpNo As Variant

    Dim answer As VbMsgBoxResult

    Dim i As Long
    Dim dictKey As Variant

    ' Display the message box with Yes and No buttons
    answer = MsgBox("Do you want to use previous calculated JSON text file", vbYesNo + vbQuestion, "JSON text file")

    ' Check which button the user clicked
    If answer = vbYes Then
        ' Code to execute if user clicks Yes
        ' MsgBox "User clicked Yes for JSON"

        jsonPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", jsonPath)  ' JSON file path
        If Not UBound(jsonPathArr) = 1 Then
            MsgBox "Please select only one JSON file"
            Exit Function
        End If

        Set allUpClause8UseAsMushakOrBillOfEntryDic = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", jsonPathArr(1))

        ' Display the message box with Yes and No buttons
        answer = MsgBox("Do you want to use UP file with previous calculated JSON text file", vbYesNo + vbQuestion, "UP file")

        ' Check which button the user clicked
        If answer = vbYes Then
            ' Code to execute if user clicks Yes
            ' MsgBox "User clicked Yes for UP file"

            upPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", initialUpPath)

            For i = LBound(upPathArr) To UBound(upPathArr) ' create dictionary as mushak or bill of entry

                Application.DisplayAlerts = False
                Set currentUpWb = Workbooks.Open(upPathArr(i))
                Set currentUpWs = currentUpWb.Worksheets(2)

                curentUpNo = Application.Run("helperFunctionGetData.upNoFromProvidedWs", currentUpWs)
                Set curentUpClause8Dict = Application.Run("general_utility_functions.upClause8InformationFromProvidedWs", currentUpWs)

                currentUpWb.Close SaveChanges:=False
                Application.DisplayAlerts = True

                For Each dictKey In curentUpClause8Dict.keys

                    If Not allUpClause8UseAsMushakOrBillOfEntryDic.Exists(dictKey) Then ' create mushak or bill of entry dictionary

                        allUpClause8UseAsMushakOrBillOfEntryDic.Add dictKey, CreateObject("Scripting.Dictionary")

                    End If

                    If Not allUpClause8UseAsMushakOrBillOfEntryDic(dictKey).Exists(curentUpNo) Then ' create current UP dictionary as inner dictionary of mushak or bill of entry dictionary

                        allUpClause8UseAsMushakOrBillOfEntryDic(dictKey).Add curentUpNo, CreateObject("Scripting.Dictionary")

                            ' Individual used Qty. and value keep in inner UP dictionary
                        allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)(curentUpNo)("inThisUpUsedQtyOfGoods") = curentUpClause8Dict(dictKey)("inThisUpUsedQtyOfGoods")
                        allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)(curentUpNo)("inThisUpUsedValueOfGoods") = curentUpClause8Dict(dictKey)("inThisUpUsedValueOfGoods")

                            ' Sum of all UP used Qty. and value keep at mushak or bill of entry dictionary only first time. There is no chance to repeat sum, if any UP select again.
                        allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("sumOfAllUpUsedQty") = allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("sumOfAllUpUsedQty") + curentUpClause8Dict(dictKey)("inThisUpUsedQtyOfGoods")
                        allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("sumOfAllUpUsedValue") = allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("sumOfAllUpUsedValue") + curentUpClause8Dict(dictKey)("inThisUpUsedValueOfGoods")
                            ' Concate all UP 
                        allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("usedUpList") = allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("usedUpList") & "," & curentUpNo

                            ' Same UP no. multiple time reassign but include all calculated UP
                        allUpClause8UseAsMushakOrBillOfEntryDic("allCalculatedUpList")(curentUpNo) = curentUpNo

                    End If

                Next dictKey

            Next i


        ElseIf answer = vbNo Then
            ' Code to execute if user clicks No
            ' MsgBox "User clicked No for UP file"

            Set sumUsedQtyAndValueAsMushakOrBillOfEntryFromSelectedUpFile = allUpClause8UseAsMushakOrBillOfEntryDic
            Exit Function

        End If

    ElseIf answer = vbNo Then
        ' Code to execute if user clicks No
        ' MsgBox "User clicked No for JSON"

        upPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", initialUpPath)  ' UP file path should be dynamic

            ' This inner dictionary create one time only when create brand new JSON test file, next time populate only
        allUpClause8UseAsMushakOrBillOfEntryDic.Add "allCalculatedUpList", CreateObject("Scripting.Dictionary")

        For i = LBound(upPathArr) To UBound(upPathArr) ' create dictionary as mushak or bill of entry

            Application.DisplayAlerts = False
            Set currentUpWb = Workbooks.Open(upPathArr(i))
            Set currentUpWs = currentUpWb.Worksheets(2)

            curentUpNo = Application.Run("helperFunctionGetData.upNoFromProvidedWs", currentUpWs)
            Set curentUpClause8Dict = Application.Run("general_utility_functions.upClause8InformationFromProvidedWs", currentUpWs)

            currentUpWb.Close SaveChanges:=False
            Application.DisplayAlerts = True
    
            For Each dictKey In curentUpClause8Dict.keys

                If Not allUpClause8UseAsMushakOrBillOfEntryDic.Exists(dictKey) Then ' create mushak or bill of entry dictionary

                    allUpClause8UseAsMushakOrBillOfEntryDic.Add dictKey, CreateObject("Scripting.Dictionary")

                End If

                If Not allUpClause8UseAsMushakOrBillOfEntryDic(dictKey).Exists(curentUpNo) Then ' create current UP dictionary as inner dictionary of mushak or bill of entry dictionary

                    allUpClause8UseAsMushakOrBillOfEntryDic(dictKey).Add curentUpNo, CreateObject("Scripting.Dictionary")

                        ' Individual used Qty. and value keep in inner UP dictionary
                    allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)(curentUpNo)("inThisUpUsedQtyOfGoods") = curentUpClause8Dict(dictKey)("inThisUpUsedQtyOfGoods")
                    allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)(curentUpNo)("inThisUpUsedValueOfGoods") = curentUpClause8Dict(dictKey)("inThisUpUsedValueOfGoods")

                        ' Sum of all UP used Qty. and value keep at mushak or bill of entry dictionary only first time. There is no chance to repeat sum, if any UP select again.
                    allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("sumOfAllUpUsedQty") = allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("sumOfAllUpUsedQty") + curentUpClause8Dict(dictKey)("inThisUpUsedQtyOfGoods")
                    allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("sumOfAllUpUsedValue") = allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("sumOfAllUpUsedValue") + curentUpClause8Dict(dictKey)("inThisUpUsedValueOfGoods")
                        ' Concate all UP 
                    allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("usedUpList") = allUpClause8UseAsMushakOrBillOfEntryDic(dictKey)("usedUpList") & "," & curentUpNo

                        ' Same UP no. multiple time reassign but include all calculated UP
                    allUpClause8UseAsMushakOrBillOfEntryDic("allCalculatedUpList")(curentUpNo) = curentUpNo


                End If

            Next dictKey

        Next i

    End If

    Dim sortedAllCalculatedUp As Variant
    sortedAllCalculatedUp = Application.Run("Sorting_Algorithms.upSort", allUpClause8UseAsMushakOrBillOfEntryDic("allCalculatedUpList").Keys)

    Application.Run "JsonUtilityFunction.SaveDictionaryToJsonTextFile", allUpClause8UseAsMushakOrBillOfEntryDic, jsonPath & Application.PathSeparator & _
    "UP-" & Replace(sortedAllCalculatedUp(LBound(sortedAllCalculatedUp)), "/", "-") & "-to-" & Replace(sortedAllCalculatedUp(UBound(sortedAllCalculatedUp)), "/", "-") & "-used-details-as-mushak-or-bill-of-entry" & ".json"

    Application.ScreenUpdating = True

    Set sumUsedQtyAndValueAsMushakOrBillOfEntryFromSelectedUpFile = allUpClause8UseAsMushakOrBillOfEntryDic

End Function

Private Function ExcludeElements(arr1 As Variant, arr2 As Variant) As Variant
    'exclude all the elements from first array which elements exist in second array
    Dim i As Long
    Dim j As Long

    Dim arr2Dictionary As Object
    Set arr2Dictionary = CreateObject("Scripting.Dictionary")

    Dim excludedDictionary As Object
    Set excludedDictionary = CreateObject("Scripting.Dictionary")
        
    ' Loop through the elements of arr2
    For i = LBound(arr2) To UBound(arr2)
        
        arr2Dictionary(arr2(i)) = arr2(i)
        
    Next i

    ' Loop through the elements of arr1
    For j = LBound(arr1) To UBound(arr1)

        If Not arr2Dictionary.Exists(arr1(j)) Then
            excludedDictionary(arr1(j)) = arr1(j)
        End If
        
    Next j
        
    ' Return the result array
    ExcludeElements = excludedDictionary.keys
        
End Function
   