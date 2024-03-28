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

