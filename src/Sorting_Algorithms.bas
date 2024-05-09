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