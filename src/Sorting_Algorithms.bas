Attribute VB_Name = "Sorting_Algorithms"
Option Explicit

Private Function BubbleSort(arr As Variant) As Variant

  Dim i, j As Long
  Dim n As Long
  Dim swapped As Boolean
  Dim temp As Variant

  n = UBound(arr) - LBound(arr) + 1 ' Get the number of elements in the array

  ' Outer loop to traverse through all array elements
  For i = 1 To n - 1
    swapped = False ' Flag to check if any swaps occurred in the inner loop
    ' Inner loop to compare adjacent elements
    For j = LBound(arr) To UBound(arr) - i
      If arr(j) > arr(j + 1) Then
        ' Swap elements if they are in the wrong order
        temp = arr(j)
        arr(j) = arr(j + 1)
        arr(j + 1) = temp
        swapped = True
      End If
    Next j
    ' If no swaps occurred in the inner loop, the array is already sorted
    If Not swapped Then
      Exit For
    End If

  Next i

  BubbleSort = arr ' Return the sorted array

End Function

Private Function upSort(upArr As Variant) As Variant
  ' This function sort UP array

  Dim yearMultiplyKeyDict As Object
  Set yearMultiplyKeyDict = CreateObject("Scripting.Dictionary")

  Dim sortedKeys As Variant
  Dim sortedUp As Variant

  Dim yearMultiplyKey As Variant
  Dim extractedUpAndUpYear As Object
  Dim i, j As Long

  For i = LBound(upArr) To UBound(upArr)

    Set extractedUpAndUpYear = Application.Run("general_utility_functions.upNoAndYearExtracAsDict", upArr(i))
    yearMultiplyKey = extractedUpAndUpYear("only_up_year") * extractedUpAndUpYear("only_up_year") * extractedUpAndUpYear("only_up_year")
    yearMultiplyKeyDict(extractedUpAndUpYear("only_up_no") + yearMultiplyKey) = upArr(i)

  Next i

  sortedKeys = Application.Run("Sorting_Algorithms.BubbleSort", yearMultiplyKeyDict.Keys)

  ReDim sortedUp(LBound(sortedKeys) To UBound(sortedKeys))

  For j = LBound(sortedKeys) To UBound(sortedKeys)

    sortedUp(j) = yearMultiplyKeyDict(sortedKeys(j))

  Next j

  upSort = sortedUp

End Function

Function SplitSequence(arr As Variant) As Object
    'this function received an numbers array. And return a dictionary
    'split every sequence and add to inner dictionary
    'every inner dictionary have "sequenceStart" and "sequenceEnd" keys
    'if have any single sequence "sequenceStart" and "sequenceEnd" keys hold the same number
    
    Dim resultDict As Object
    Set resultDict = CreateObject("Scripting.Dictionary")

    Dim sequenceStart As Variant
    Dim sequenceEnd As Variant
    Dim i As Long
    
    ' Initialize the start of the sequence
    sequenceStart = arr(LBound(arr))
    sequenceEnd = sequenceStart

    ' Loop through the array
    For i = LBound(arr) + 1 To UBound(arr)
      ' Check if the current number is sequential
      If arr(i) = sequenceEnd + 1 Then
        ' Update the end of the sequence
        sequenceEnd = arr(i)
      Else
      ' Add the current sequence to the result dictionary
      If Not resultDict.Exists(resultDict.Count + 1) Then ' create inner sequence dictionary
        resultDict.Add resultDict.Count + 1, CreateObject("Scripting.Dictionary")
      End If
        'already add inner dictionary so, resultDict.Count point to same inner dictionary
        resultDict(resultDict.Count)("sequenceStart") = sequenceStart
        resultDict(resultDict.Count)("sequenceEnd") = sequenceEnd

        ' Start a new sequence
        sequenceStart = arr(i)
        sequenceEnd = sequenceStart
      End If
    Next i

    ' Add the last sequence to the result dictionary
    If Not resultDict.Exists(resultDict.Count + 1) Then ' create sequence dictionary
        resultDict.Add resultDict.Count + 1, CreateObject("Scripting.Dictionary")
    End If

    'already add inner dictionary so, resultDict.Count point to same inner dictionary
    resultDict(resultDict.Count)("sequenceStart") = sequenceStart
    resultDict(resultDict.Count)("sequenceEnd") = sequenceEnd

  Set SplitSequence = resultDict
    
End Function

Function SplituPSequence(upArr As Variant) As Object

  'this function received an UP-Numbers array. And return a dictionary
  'sort every UP ascending order
  'split every UP sequence and add to inner dictionary
  'every inner dictionary have "sequenceStart" and "sequenceEnd" keys
  'if have any single sequence "sequenceStart" and "sequenceEnd" keys hold the same UP number
    
  Dim yearMultiplyKeyDict As Object
  Set yearMultiplyKeyDict = CreateObject("Scripting.Dictionary")

  Dim sortedKeys As Variant
  Dim splitedKeysDict As Object

  Dim yearMultiplyKey As Variant
  Dim extractedUpAndUpYear As Object
  Dim i As Long

  For i = LBound(upArr) To UBound(upArr)

    Set extractedUpAndUpYear = Application.Run("general_utility_functions.upNoAndYearExtracAsDict", upArr(i))
    yearMultiplyKey = extractedUpAndUpYear("only_up_year") * extractedUpAndUpYear("only_up_year") * extractedUpAndUpYear("only_up_year")
    yearMultiplyKeyDict(extractedUpAndUpYear("only_up_no") + yearMultiplyKey) = upArr(i)

  Next i

  sortedKeys = Application.Run("Sorting_Algorithms.BubbleSort", yearMultiplyKeyDict.Keys) 'sort ascending order

  Set splitedKeysDict = Application.Run("Sorting_Algorithms.SplitSequence", sortedKeys) 'split every sequence

  Dim dictKey As Variant

  For Each dictKey In splitedKeysDict.keys

    splitedKeysDict(dictKey)("sequenceStart") = yearMultiplyKeyDict(splitedKeysDict(dictKey)("sequenceStart")) 'pick UP no.
    splitedKeysDict(dictKey)("sequenceEnd") = yearMultiplyKeyDict(splitedKeysDict(dictKey)("sequenceEnd")) 'pick UP no.

  Next dictKey

  Set SplituPSequence = splitedKeysDict
    
End Function

Private Function FindMaxTwoNumbers(arr As Variant) As Object

  Dim maxTwo As Object
  Set maxTwo = CreateObject("Scripting.Dictionary")

  If UBound(arr) - LBound(arr) < 1 Then
    MsgBox "Must be array have two or more elements"
    Set FindMaxTwoNumbers = maxTwo
    Exit Function
  End If

  Dim sortedAscendingArr As Variant

  sortedAscendingArr = Application.Run("Sorting_Algorithms.BubbleSort", arr) 'sort ascending order

  maxTwo("firstMax") = sortedAscendingArr(UBound(sortedAscendingArr))
  maxTwo("secondMax") = sortedAscendingArr(UBound(sortedAscendingArr) - 1)

  Set FindMaxTwoNumbers = maxTwo

End Function