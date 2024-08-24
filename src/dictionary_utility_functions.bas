Attribute VB_Name = "dictionary_utility_functions"
Option Explicit

'Private Function CreateDictionary()
'    Dim dictionary As Object
'    Set dictionary = CreateObject("Scripting.Dictionary")
'    Set CreateDictionary = dictionary
'End Function

Private Function CreateDicWithProvidedKeysAndValues(keysArray As Variant, valuesArray As Variant) As Object

    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    Dim removedAllInvalidChrFromKeys As Variant

    Dim i As Long

    ' Add keys with values
    For i = LBound(keysArray) To UBound(keysArray)

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", keysArray(i))   'remove all invalid characters for use dic keys
        
        If dictionary.Exists(removedAllInvalidChrFromKeys) Then
            MsgBox "Dictionary Key """ & removedAllInvalidChrFromKeys & """ Already Exists"
            Exit Function
        Else
            dictionary(removedAllInvalidChrFromKeys) = valuesArray(i)
        End If

    Next i

    ' Return the created dictionary
    Set CreateDicWithProvidedKeysAndValues = dictionary

End Function


Private Function AddKeysWithPrimary(dictionary As Object, primaryKey As Variant, keysArray As Variant) As Object

    Dim removedAllInvalidChrFromKeys As Variant
    Dim i As Long

    ' Add new keys with primary value
    For i = LBound(keysArray) To UBound(keysArray)

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", keysArray(i))   'remove all invalid characters for use dic keys
        dictionary(removedAllInvalidChrFromKeys) = primaryKey

    Next i

    ' Return the modified dictionary
    Set AddKeysWithPrimary = dictionary
End Function


Private Function AddKeysAndValueSame(dictionary As Object, keysArray As Variant) As Object
        
    Dim removedAllInvalidChrFromKeys As Variant
        
    Dim i As Long
    ' Add same keys and value
    For i = LBound(keysArray) To UBound(keysArray)
        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", keysArray(i))   'remove all invalid characters for use dic keys
        dictionary(removedAllInvalidChrFromKeys) = keysArray(i)
    Next i

    ' Return the modified dictionary
    Set AddKeysAndValueSame = dictionary
End Function


Private Function addKeysAndValueToDic(dictionary As Object, key As Variant, value As Variant) As Object

    Dim removedAllInvalidChrFromKeys As Variant

    ' Add key with values

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", key)   'remove all invalid characters for use dic keys
        
        If dictionary.Exists(removedAllInvalidChrFromKeys) Then
            MsgBox "Dictionary Key """ & removedAllInvalidChrFromKeys & """ Already Exists"
            Exit Function
        Else
            dictionary.Add removedAllInvalidChrFromKeys, value
        End If

    ' Return the modified dictionary
    Set addKeysAndValueToDic = dictionary

End Function

 
Private Function PutDictionaryValuesIntoWorksheet(wsRange As Range, dict As Object, keysPrint As Boolean, itemsPrint As Boolean, printOnColumn As Boolean)
    ' wsRange is just starting one cell address, the function dynamically resizes the range
    
    If dict.Count > 0 Then
    
        If (keysPrint And itemsPrint And printOnColumn) Then
    
            wsRange.Resize(dict.Count, 1).value = Application.Run("general_utility_functions.oneDArrayConvertToTwoDArray", dict.keys)
            wsRange.Offset(0, 1).Resize(dict.Count, 1).value = Application.Run("general_utility_functions.oneDArrayConvertToTwoDArray", dict.items)
            
        ElseIf (keysPrint And printOnColumn) Then
    
            wsRange.Resize(dict.Count, 1).value = Application.Run("general_utility_functions.oneDArrayConvertToTwoDArray", dict.keys)
    
        ElseIf (itemsPrint And printOnColumn) Then
    
            wsRange.Resize(dict.Count, 1).value = Application.Run("general_utility_functions.oneDArrayConvertToTwoDArray", dict.items)
    
        ElseIf (keysPrint And itemsPrint) Then
    
            wsRange.Resize(1, dict.Count).value = dict.keys
            wsRange.Offset(1, 0).Resize(1, dict.Count).value = dict.items
    
        ElseIf (keysPrint) Then
    
            wsRange.Resize(1, dict.Count).value = dict.keys
    
        ElseIf (itemsPrint) Then
    
            wsRange.Resize(1, dict.Count).value = dict.items
    
        End If
    
    End If

End Function


Private Function CreateMushakOrBillOfEntryDbDict(dict As Object, mushakOrBillOfEntrySourceArr As Variant, lcCol As Integer, mushakOrBillOfEntryCol As Integer, qtyCol As Integer, valueCol As Integer, discriptionCol As Integer, propertiesArr As Variant, propertiesColsArr As Variant) As Object

    Dim tempMuOrBillKey As Variant
    Dim propertiesValArr As Variant

    ReDim propertiesValArr(LBound(propertiesColsArr) To UBound(propertiesColsArr))

    Dim tempMuOrBillKeyDic As Object

    Dim i As Long
    Dim j As Long

    For i = LBound(mushakOrBillOfEntrySourceArr) To UBound(mushakOrBillOfEntrySourceArr)

        For j = LBound(propertiesValArr) To UBound(propertiesValArr)

            propertiesValArr(j) = mushakOrBillOfEntrySourceArr(i, propertiesColsArr(j))
            
        Next j

        tempMuOrBillKey = Application.Run("general_utility_functions.dictKeyGeneratorWithLcMushakOrBillOfEntryQtyAndValue", mushakOrBillOfEntrySourceArr(i, lcCol), mushakOrBillOfEntrySourceArr(i, mushakOrBillOfEntryCol), mushakOrBillOfEntrySourceArr(i, qtyCol), mushakOrBillOfEntrySourceArr(i, valueCol))

        Set tempMuOrBillKeyDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)
        If dict.Exists(tempMuOrBillKey) Then
            
            dict(tempMuOrBillKey)("Entry_Count") = dict(tempMuOrBillKey)("Entry_Count") + 1
            dict(tempMuOrBillKey)("All_Entry_Discription") = dict(tempMuOrBillKey)("All_Entry_Discription") & Chr(10) & mushakOrBillOfEntrySourceArr(i, discriptionCol)
            
        Else
        
            dict.Add tempMuOrBillKey, tempMuOrBillKeyDic
            dict(tempMuOrBillKey)("Entry_Count") = 1
            dict(tempMuOrBillKey)("All_Entry_Discription") = mushakOrBillOfEntrySourceArr(i, discriptionCol)
            
        End If
        

    Next i
    
    Set CreateMushakOrBillOfEntryDbDict = dict
    
End Function


Private Function SortDictionaryByKey(dict As Object _
                  , Optional sortorder As XlSortOrder = xlAscending) As Object
    
    Dim arrList As Object
    Set arrList = CreateObject("System.Collections.ArrayList")
    
    ' Put keys in an ArrayList
    Dim key As Variant, coll As New Collection
    For Each key In dict
        arrList.Add key
    Next key
    
    ' Sort the keys
    arrList.Sort
    
    ' For descending order, reverse
    If sortorder = xlDescending Then
        arrList.Reverse
    End If
    
    ' Create new dictionary
    Dim dictNew As Object
    Set dictNew = CreateObject("Scripting.Dictionary")
    
    ' Read through the sorted keys and add to new dictionary
    For Each key In arrList
        dictNew.Add key, dict(key)
    Next key
    
    ' Clean up
    Set arrList = Nothing
    Set dict = Nothing
    
    ' Return the new dictionary
    Set SortDictionaryByKey = dictNew
        
End Function

Private Function mergeDict(mainDict As Object, addingDict As Object) As Object
    'this function received two dictionaries and merge them, then return merged dictionary

    Dim dictKey As Variant
    Dim i As Long
    For i = 0 To addingDict.Count - 1

        dictKey = addingDict.keys()(i)

        If mainDict.Exists(dictKey) Then
            MsgBox "Dictionary Key """ & dictKey & """ Already Exists"
            Exit Function
        Else
            mainDict.Add dictKey, addingDict(dictKey)
        End If

    Next i

    Set mergeDict = mainDict

End Function


Private Function sumOfProvidedKeys(dict As Object, arrOfKeys As Variant) As Variant
    'this function received a dictionary and a array of keys then
    ' sum of provided key's value and return the sum

    Dim element  As Variant
    Dim removedAllInvalidChrFromKeys As Variant
    Dim sum As Variant
    sum = 0

    For Each element In arrOfKeys

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", element)    'remove all invalid characters for use dic keys

        If dict.Exists(removedAllInvalidChrFromKeys) Then

            sum = sum + dict(removedAllInvalidChrFromKeys)

        Else

            MsgBox "Dictionary Key """ & removedAllInvalidChrFromKeys & """ Not Found"
            Exit Function

        End If

    Next

    sumOfProvidedKeys = sum

End Function

Private Function sumOfInnerDictOfProvidedKeys(dict As Object, arrOfKeys As Variant) As Variant
    'received a one level nested dictionary and a array of keys then
    ' sum of all inner dictionary of provided key's value and return the sum

    Dim sum As Variant
    sum = 0

    Dim dicKey As Variant

    For Each dicKey In dict.keys
        
        sum = sum + Application.Run("dictionary_utility_functions.sumOfProvidedKeys", dict(dicKey), arrOfKeys)

    Next dicKey

    sumOfInnerDictOfProvidedKeys = sum

End Function

Private Function arrSpecificColumnGroupAndSpecificColumnSumAsGroup(srcArr As Variant, columnOfGroup As Integer, columnOfSum As Integer) As Object
        
    Dim removedAllInvalidChrFromKeys As Variant

    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")

    Dim i As Long
    ' Group as same keys
    For i = LBound(srcArr) To UBound(srcArr)
        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", srcArr(i, columnOfGroup))   'remove all invalid characters for use dic keys
        ' sum as group
        dictionary(removedAllInvalidChrFromKeys) = dictionary(removedAllInvalidChrFromKeys) + srcArr(i, columnOfSum)
    Next i

    ' Return the dictionary
    Set arrSpecificColumnGroupAndSpecificColumnSumAsGroup = dictionary

End Function