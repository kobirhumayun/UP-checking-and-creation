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
