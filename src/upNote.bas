Attribute VB_Name = "upNote"
Option Explicit

Private Function putUpSummary(noteWorksheet As Worksheet, sourceDataAsDicUpIssuingStatus As Object, upClause8InfoClassifiedPartDic As Object, newUp As String)

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim lcCountRow As Long
    lcCountRow = noteWorksheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("exportLcSalesContractBengaliTxt"), LookAt:=xlPart).Row

    noteWorksheet.Range("C" & lcCountRow - 1).value = vsCodeNotSupportedOrBengaliTxtDictionary("upAppNoPart1BengaliTxt") _
        & newUp & vsCodeNotSupportedOrBengaliTxtDictionary("upAppNoPart2BengaliTxt")
    noteWorksheet.Range("F" & lcCountRow).value = sourceDataAsDicUpIssuingStatus.Count

    Dim totalUsedQtyOfGoods, totalUsedValueOfGoods, totalUsedQtyOfYarn As Variant

    totalUsedQtyOfGoods = upClause8InfoClassifiedPartDic("yarnImportQty") + _
        upClause8InfoClassifiedPartDic("yarnLocalQty") + _
        upClause8InfoClassifiedPartDic("dyesQty") + _
        upClause8InfoClassifiedPartDic("stretchWrappingFilmQty") + _
        upClause8InfoClassifiedPartDic("chemicalsImportQty") + _
        upClause8InfoClassifiedPartDic("chemicalsLocalQty")

    totalUsedValueOfGoods = upClause8InfoClassifiedPartDic("yarnImportValue") + _
        upClause8InfoClassifiedPartDic("yarnLocalValue") + _
        upClause8InfoClassifiedPartDic("dyesValue") + _
        upClause8InfoClassifiedPartDic("stretchWrappingFilmValue") + _
        upClause8InfoClassifiedPartDic("chemicalsImportValue") + _
        upClause8InfoClassifiedPartDic("chemicalsLocalValue")

    totalUsedQtyOfYarn = upClause8InfoClassifiedPartDic("yarnImportQty") + upClause8InfoClassifiedPartDic("yarnLocalQty")

    noteWorksheet.Range("F" & lcCountRow + 1).value = totalUsedQtyOfGoods
    noteWorksheet.Range("F" & lcCountRow + 2).value = totalUsedValueOfGoods
    noteWorksheet.Range("F" & lcCountRow + 3).value = totalUsedQtyOfYarn
    
    Dim dicKey As Variant
    Dim exportValue, exportQty As Variant

    exportValue = 0
    exportQty = 0

    For Each dicKey In sourceDataAsDicUpIssuingStatus.keys

            exportValue = exportValue + Application.Run("createUp.valueInUsd", sourceDataAsDicUpIssuingStatus(dicKey))
            exportQty = exportQty + Application.Run("createUp.qtyInYds", sourceDataAsDicUpIssuingStatus(dicKey))

    Next dicKey

        'clear first, because when manual ref. from UP sheet it's an array & withous clear error occur
    Range("K" & lcCountRow & ":L" & lcCountRow + 1).ClearContents

    noteWorksheet.Range("K" & lcCountRow).value = exportValue
    noteWorksheet.Range("K" & lcCountRow + 1).value = exportQty
    noteWorksheet.Range("K" & lcCountRow + 2).value = (exportValue - totalUsedValueOfGoods) / totalUsedValueOfGoods * 100
    noteWorksheet.Range("K" & lcCountRow + 3).value = upClause8InfoClassifiedPartDic("dyesQty") + _
        upClause8InfoClassifiedPartDic("stretchWrappingFilmQty") + _
        upClause8InfoClassifiedPartDic("chemicalsImportQty") + _
        upClause8InfoClassifiedPartDic("chemicalsLocalQty")



End Function

Private Function putLcInfo(noteWorksheet As Worksheet, sourceDataAsDicUpIssuingStatus As Object)

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim topRow, bottomRow As Long

    topRow = noteWorksheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("bbLcScNoAndDtBengaliTxt"), LookAt:=xlPart).Row
    bottomRow = noteWorksheet.Range("D" & topRow).End(xlDown).Row

        'one row down to take LC info only
    topRow = topRow + 1

    Dim workingRange As Range
    Set workingRange = noteWorksheet.Range("A" & topRow & ":" & "M" & bottomRow)

            'keep only one LC
        If workingRange.Rows.Count > 1 Then

            workingRange.Rows("2:" & workingRange.Rows.Count).EntireRow.Delete

        End If

        'insert rows as lc count, note already one row exist
        If sourceDataAsDicUpIssuingStatus.Count > 1 Then

            Dim i As Long
            For i = 1 To sourceDataAsDicUpIssuingStatus.Count - 1
                workingRange.Rows("2").EntireRow.Insert
            Next i

        End If

        Set workingRange = workingRange.Resize(sourceDataAsDicUpIssuingStatus.Count)
        workingRange.Clear
        Application.Run "utility_formating_fun.rangeFormat", workingRange, "Calibri", 10, False, True, xlCenter, xlCenter, "General"

        Dim j As Long
        Dim dicKey As Variant

        For j = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

            dicKey = sourceDataAsDicUpIssuingStatus.keys()(j)

            workingRange.Range("C" & j + 1).value = j + 1
            workingRange.Range("D" & j + 1).value = Application.Run("createUp.combinLcAndAmnd", sourceDataAsDicUpIssuingStatus(dicKey))
            workingRange.Range("E" & j + 1).value = Application.Run("createUp.valueInUsd", sourceDataAsDicUpIssuingStatus(dicKey))
            workingRange.Range("F" & j + 1).value = Application.Run("createUp.qtyInYds", sourceDataAsDicUpIssuingStatus(dicKey))
            workingRange.Range("G" & j + 1).value = sourceDataAsDicUpIssuingStatus(dicKey)("ShipmentDate")
            workingRange.Range("G" & j + 1 & ":H" & j + 1).Merge
            workingRange.Range("I" & j + 1).value = sourceDataAsDicUpIssuingStatus(dicKey)("ExpiryDate")
            workingRange.Range("I" & j + 1 & ":J" & j + 1).Merge
            workingRange.Range("K" & j + 1).value = Application.Run("createUp.combinUdIpExpMlc", sourceDataAsDicUpIssuingStatus(dicKey))
            workingRange.Range("K" & j + 1 & ":M" & j + 1).Merge

        Next j

        workingRange.Range("E1:F" & workingRange.Rows.Count).Style = "Comma"
        Application.Run "utility_formating_fun.SetBorderThin", workingRange.Range("C1:M" & workingRange.Rows.Count)

End Function

Private Function putUdIpExpInfo(noteWorksheet As Worksheet, sourceDataAsDicUpIssuingStatus As Object)

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim topRow, bottomRow As Long

    topRow = noteWorksheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("udIpExpNoAndDtBengaliTxt"), LookAt:=xlPart).Row
    bottomRow = noteWorksheet.Range("J" & topRow).End(xlDown).Row

        'one row down to take UD info only
    topRow = topRow + 1

    Dim workingRange As Range
    Set workingRange = noteWorksheet.Range("A" & topRow & ":" & "M" & bottomRow)

            'keep only one UD
        If workingRange.Rows.Count > 1 Then

            workingRange.Rows("2:" & workingRange.Rows.Count).EntireRow.Delete

        End If

        'insert rows as lc count, note already one row exist
        If sourceDataAsDicUpIssuingStatus.Count > 1 Then

            Dim i As Long
            For i = 1 To sourceDataAsDicUpIssuingStatus.Count - 1
                workingRange.Rows("2").EntireRow.Insert
            Next i

        End If

        Set workingRange = workingRange.Resize(sourceDataAsDicUpIssuingStatus.Count)
        workingRange.Clear
        Application.Run "utility_formating_fun.rangeFormat", workingRange, "Calibri", 11, False, True, xlCenter, xlCenter, "General"
        Application.Run "utility_formating_fun.rangeFormat", workingRange.Columns(10), "SutonnyMJ", 11, False, True, xlCenter, xlCenter, "General"


        Dim j, l, m As Long
        Dim dicKey As Variant
        Dim innerDicKey As Variant

        Dim tempWidthStr As Object
        Dim tempWeightStr As Object
        Dim temp As Variant


        For j = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

            dicKey = sourceDataAsDicUpIssuingStatus.keys()(j)

            workingRange.Range("C" & j + 1).value = j + 1
            workingRange.Range("D" & j + 1).value = Application.Run("createUp.combinUdIpExpAndDt", sourceDataAsDicUpIssuingStatus(dicKey))
            workingRange.Range("D" & j + 1 & ":G" & j + 1).Merge

            Set tempWidthStr = CreateObject("Scripting.Dictionary")
            Set tempWeightStr = CreateObject("Scripting.Dictionary")

            For Each innerDicKey In sourceDataAsDicUpIssuingStatus(dicKey)("consumptionRange").keys

                    'create unique width & weight
                tempWidthStr(sourceDataAsDicUpIssuingStatus(dicKey)("consumptionRange")(innerDicKey)("width").value) = Null
                tempWeightStr(sourceDataAsDicUpIssuingStatus(dicKey)("consumptionRange")(innerDicKey)("weight").value) = Null

            Next innerDicKey

                'sort
            temp = Application.Run("Sorting_Algorithms.BubbleSort", tempWidthStr.keys)

                'formate
            For l = LBound(temp) To UBound(temp)
                temp(l) = Format(temp(l), "0.00")
            Next l
            
                'put width
            workingRange.Range("H" & j + 1).value = Join(temp, ",")

                'sort
            temp = Application.Run("Sorting_Algorithms.BubbleSort", tempWeightStr.keys)
            
                'format
            For m = LBound(temp) To UBound(temp)
                temp(m) = Format(temp(m), "0.00")
            Next m
            
                'put weight
            workingRange.Range("I" & j + 1).value = Join(temp, ",")

            workingRange.Range("J" & j + 1).value = vsCodeNotSupportedOrBengaliTxtDictionary("denimFabricsBengaliTxt")
            workingRange.Range("J" & j + 1 & ":M" & j + 1).Merge

        Next j

        Application.Run "utility_formating_fun.SetBorderThin", workingRange.Range("C1:M" & workingRange.Rows.Count)

End Function

Private Function putBuyerAndBankInfo(noteWorksheet As Worksheet, sourceDataAsDicUpIssuingStatus As Object)

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim topRow, bottomRow As Long

    topRow = noteWorksheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("issuingBankNameAndAddressBengaliTxt"), LookAt:=xlPart).Row
    bottomRow = noteWorksheet.Range("C" & topRow).End(xlDown).Row

        'one row down to take buyer info only
    topRow = topRow + 1

    Dim workingRange As Range
    Set workingRange = noteWorksheet.Range("A" & topRow & ":" & "M" & bottomRow)

            'keep only one buyer
        If workingRange.Rows.Count > 1 Then

            workingRange.Rows("2:" & workingRange.Rows.Count).EntireRow.Delete

        End If

        'insert rows as lc count, note already one row exist
        If sourceDataAsDicUpIssuingStatus.Count > 1 Then

            Dim i As Long
            For i = 1 To sourceDataAsDicUpIssuingStatus.Count - 1
                workingRange.Rows("2").EntireRow.Insert
            Next i

        End If

        Set workingRange = workingRange.Resize(sourceDataAsDicUpIssuingStatus.Count)
        workingRange.Clear
        Application.Run "utility_formating_fun.rangeFormat", workingRange, "Calibri", 10, False, True, xlCenter, xlCenter, "General"

        Dim j As Long
        Dim dicKey As Variant

        For j = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

            dicKey = sourceDataAsDicUpIssuingStatus.keys()(j)

            workingRange.Range("C" & j + 1).value = j + 1
            workingRange.Range("D" & j + 1).value = sourceDataAsDicUpIssuingStatus(dicKey)("LCIssuingBank")
            workingRange.Range("D" & j + 1 & ":F" & j + 1).Merge
            workingRange.Range("G" & j + 1).value = sourceDataAsDicUpIssuingStatus(dicKey)("NameofBuyers")
            workingRange.Range("G" & j + 1 & ":L" & j + 1).Merge

        Next j

        Application.Run "utility_formating_fun.SetBorderThin", workingRange.Range("C1:L" & workingRange.Rows.Count)

End Function

Private Function putVerifiedInfo(noteWorksheet As Worksheet, sourceDataAsDicUpIssuingStatus As Object)

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim topRow, bottomRow As Long

    topRow = noteWorksheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("submittedInfoBengaliTxt"), LookAt:=xlPart).Row + 1
    bottomRow = noteWorksheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("rawMaterialNameAndDescriptionBengaliTxt"), LookAt:=xlPart).Row - 3

    Dim workingRange As Range
    Set workingRange = noteWorksheet.Range("A" & topRow & ":" & "M" & bottomRow)

    'keep only one row
    workingRange.Rows("2:" & workingRange.Rows.Count).EntireRow.Delete

    Dim i As Long
    Dim totalInsertedRows As Long

    'insert rows as lc count, note already one row exist
    If sourceDataAsDicUpIssuingStatus.Count > 1 Then

        totalInsertedRows = sourceDataAsDicUpIssuingStatus.Count * 9 + sourceDataAsDicUpIssuingStatus.Count - 2

        For i = 1 To totalInsertedRows
            workingRange.Rows("2").EntireRow.Insert
        Next i

    Else

        totalInsertedRows = 8

        For i = 1 To totalInsertedRows
            workingRange.Rows("2").EntireRow.Insert
        Next i

    End If

    Set workingRange = workingRange.Resize(totalInsertedRows + 1)
    workingRange.Clear
    Application.Run "utility_formating_fun.rangeFormat", workingRange, "Calibri", 11, False, True, xlCenter, xlCenter, "General"
    Application.Run "utility_formating_fun.rangeFormat", workingRange.Columns(4), "SutonnyMJ", 11, False, True, xlCenter, xlCenter, "General"
    Application.Run "utility_formating_fun.rangeFormat", workingRange.Columns(13), "SutonnyMJ", 11, False, True, xlCenter, xlCenter, "General"

    Dim j As Long
    Dim rowTracker As Long
    Dim dicKey As Variant

    rowTracker = 0

    For j = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

        dicKey = sourceDataAsDicUpIssuingStatus.keys()(j)


        workingRange.Range("B" & rowTracker + 1).value = Chr(j + 97) & ")"
        workingRange.Range("C" & rowTracker + 1).value = "1)"
        workingRange.Range("D" & rowTracker + 1).value = vsCodeNotSupportedOrBengaliTxtDictionary("buyerNameBengaliTxt")
        workingRange.Range("E" & rowTracker + 1).value = sourceDataAsDicUpIssuingStatus(dicKey)("NameofBuyers")
        workingRange.Range("E" & rowTracker + 1 & ":G" & rowTracker + 1).Merge
        workingRange.Range("H" & rowTracker + 1).value = sourceDataAsDicUpIssuingStatus(dicKey)("NameofBuyers")
        workingRange.Range("H" & rowTracker + 1 & ":L" & rowTracker + 1).Merge
        workingRange.Range("M" & rowTracker + 1).value = vsCodeNotSupportedOrBengaliTxtDictionary("foundCorrectBengaliTxt")

        workingRange.Range("C" & rowTracker + 2).value = "2)"
        workingRange.Range("D" & rowTracker + 2).value = vsCodeNotSupportedOrBengaliTxtDictionary("udNoBengaliTxt")
        workingRange.Range("E" & rowTracker + 2).value = Application.Run("createUp.combinUdIpExpAndDt", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("E" & rowTracker + 2 & ":G" & rowTracker + 2).Merge
        workingRange.Range("H" & rowTracker + 2).value = Application.Run("createUp.combinUdIpExpAndDt", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("H" & rowTracker + 2 & ":L" & rowTracker + 2).Merge
        workingRange.Range("M" & rowTracker + 2).value = vsCodeNotSupportedOrBengaliTxtDictionary("foundCorrectBengaliTxt")

        workingRange.Range("C" & rowTracker + 3).value = "3)"
        workingRange.Range("D" & rowTracker + 3).value = vsCodeNotSupportedOrBengaliTxtDictionary("mLcExpIpNoBengaliTxt")
        workingRange.Range("E" & rowTracker + 3).value = Application.Run("createUp.combinUdIpExpMlc", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("E" & rowTracker + 3 & ":G" & rowTracker + 3).Merge
        workingRange.Range("H" & rowTracker + 3).value = Application.Run("createUp.combinUdIpExpMlc", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("H" & rowTracker + 3 & ":L" & rowTracker + 3).Merge
        workingRange.Range("M" & rowTracker + 3).value = vsCodeNotSupportedOrBengaliTxtDictionary("foundCorrectBengaliTxt")

        workingRange.Range("C" & rowTracker + 4).value = "4)"
        workingRange.Range("D" & rowTracker + 4).value = vsCodeNotSupportedOrBengaliTxtDictionary("sellerNameBengaliTxt")
        workingRange.Range("E" & rowTracker + 4).value = "Pioneer Denim Denim Ltd"
        workingRange.Range("E" & rowTracker + 4 & ":G" & rowTracker + 4).Merge
        workingRange.Range("H" & rowTracker + 4).value = "Pioneer Denim Denim Ltd"
        workingRange.Range("H" & rowTracker + 4 & ":L" & rowTracker + 4).Merge
        workingRange.Range("M" & rowTracker + 4).value = vsCodeNotSupportedOrBengaliTxtDictionary("foundCorrectBengaliTxt")

        workingRange.Range("C" & rowTracker + 5).value = "5)"
        workingRange.Range("D" & rowTracker + 5).value = vsCodeNotSupportedOrBengaliTxtDictionary("bbLcScNoAndDtBengaliTxt")
        workingRange.Range("E" & rowTracker + 5).value = Application.Run("createUp.combinLcAndAmnd", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("E" & rowTracker + 5 & ":G" & rowTracker + 5).Merge
        workingRange.Range("H" & rowTracker + 5).value = Application.Run("createUp.combinLcAndAmnd", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("H" & rowTracker + 5 & ":L" & rowTracker + 5).Merge
        workingRange.Range("M" & rowTracker + 5).value = vsCodeNotSupportedOrBengaliTxtDictionary("foundCorrectBengaliTxt")

        workingRange.Range("C" & rowTracker + 6).value = "6)"
        workingRange.Range("D" & rowTracker + 6).value = vsCodeNotSupportedOrBengaliTxtDictionary("bbLcValueBengaliTxt")
        workingRange.Range("E" & rowTracker + 6).value = Application.Run("createUp.valueInUsd", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("E" & rowTracker + 6 & ":G" & rowTracker + 6).Merge
        workingRange.Range("H" & rowTracker + 6).value = Application.Run("createUp.valueInUsd", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("H" & rowTracker + 6 & ":L" & rowTracker + 6).Merge
        workingRange.Range("M" & rowTracker + 6).value = vsCodeNotSupportedOrBengaliTxtDictionary("foundCorrectBengaliTxt")
        workingRange.Range("E" & rowTracker + 6 & ":L" & rowTracker + 6).Style = "Comma"

        workingRange.Range("C" & rowTracker + 7).value = "7)"
        workingRange.Range("D" & rowTracker + 7).value = vsCodeNotSupportedOrBengaliTxtDictionary("qtyOfGoodsYdsBengaliTxt")
        workingRange.Range("E" & rowTracker + 7).value = Application.Run("createUp.qtyInYds", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("E" & rowTracker + 7 & ":G" & rowTracker + 7).Merge
        workingRange.Range("H" & rowTracker + 7).value = Application.Run("createUp.qtyInYds", sourceDataAsDicUpIssuingStatus(dicKey))
        workingRange.Range("H" & rowTracker + 7 & ":L" & rowTracker + 7).Merge
        workingRange.Range("M" & rowTracker + 7).value = vsCodeNotSupportedOrBengaliTxtDictionary("foundCorrectBengaliTxt")
        workingRange.Range("E" & rowTracker + 7 & ":L" & rowTracker + 7).Style = "Comma"

        workingRange.Range("C" & rowTracker + 8).value = "8)"
        workingRange.Range("D" & rowTracker + 8).value = vsCodeNotSupportedOrBengaliTxtDictionary("mLcValueBengaliTxt")
        workingRange.Range("E" & rowTracker + 8).value = "" 'actual value put manually
        workingRange.Range("E" & rowTracker + 8 & ":G" & rowTracker + 8).Merge
        workingRange.Range("H" & rowTracker + 8).value = "" 'actual value put manually
        workingRange.Range("H" & rowTracker + 8 & ":L" & rowTracker + 8).Merge
        workingRange.Range("M" & rowTracker + 8).value = vsCodeNotSupportedOrBengaliTxtDictionary("foundCorrectBengaliTxt")
        workingRange.Range("E" & rowTracker + 8 & ":L" & rowTracker + 8).Style = "Comma"

        workingRange.Range("C" & rowTracker + 9).value = "9)"
        workingRange.Range("D" & rowTracker + 9).value = vsCodeNotSupportedOrBengaliTxtDictionary("mLcValidityBengaliTxt")
        workingRange.Range("E" & rowTracker + 9).value = "" 'actual value put manually
        workingRange.Range("E" & rowTracker + 9 & ":G" & rowTracker + 9).Merge
        workingRange.Range("H" & rowTracker + 9).value = "" 'actual value put manually
        workingRange.Range("H" & rowTracker + 9 & ":L" & rowTracker + 9).Merge
        workingRange.Range("M" & rowTracker + 9).value = vsCodeNotSupportedOrBengaliTxtDictionary("foundCorrectBengaliTxt")
        workingRange.Range("E" & rowTracker + 9 & ":L" & rowTracker + 9).NumberFormat = "dd/mm/yyyy"


        Application.Run "utility_formating_fun.SetBorderThin", workingRange.Range("C" & rowTracker + 1 & ":M" & rowTracker + 9)










        rowTracker = rowTracker + 10
    Next j


End Function

Private Function putRawMaterialsQtyAsGroup(noteWorksheet As Worksheet, upClause8InfoDic As Object)

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim topRow, bottomRow As Long

    topRow = noteWorksheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("rawMaterialNameAndDescriptionBengaliTxt"), LookAt:=xlPart).Row
    bottomRow = noteWorksheet.Range("C" & topRow).End(xlDown).Row

        'one row down to take raw materials Qty. area only
    topRow = topRow + 1

    Dim workingRange As Range
    Set workingRange = noteWorksheet.Range("A" & topRow & ":" & "L" & bottomRow)

    'keep only one row
    workingRange.Rows("2:" & workingRange.Rows.Count).EntireRow.Delete

    Dim dicKey As Variant
    Dim rawMaterialsQtyGroupByGoods As Object
    Set rawMaterialsQtyGroupByGoods = CreateObject("Scripting.Dictionary")
    Dim removedAllInvalidChrFromRawMaterialsDes As String
    Dim totalUsedQtyOfGoods, totalUsedValueOfGoods As Variant

    totalUsedQtyOfGoods = 0
    totalUsedValueOfGoods = 0

    For Each dicKey In upClause8InfoDic.keys

        removedAllInvalidChrFromRawMaterialsDes = Application.Run("general_utility_functions.RemoveInvalidChars", upClause8InfoDic(dicKey)("nameOfGoods")) 'remove all invalid characters

        If Not rawMaterialsQtyGroupByGoods.Exists(removedAllInvalidChrFromRawMaterialsDes) Then ' create group by goods dictionary

            rawMaterialsQtyGroupByGoods.Add removedAllInvalidChrFromRawMaterialsDes, CreateObject("Scripting.Dictionary")
            rawMaterialsQtyGroupByGoods(removedAllInvalidChrFromRawMaterialsDes).Add "concatedHsCode", CreateObject("Scripting.Dictionary") 'for unique value

        End If

        rawMaterialsQtyGroupByGoods(removedAllInvalidChrFromRawMaterialsDes)("nameOfGoods") = upClause8InfoDic(dicKey)("nameOfGoods")

        rawMaterialsQtyGroupByGoods(removedAllInvalidChrFromRawMaterialsDes)("concatedHsCode")(upClause8InfoDic(dicKey)("hsCode")) = Null 'just for unique value

        rawMaterialsQtyGroupByGoods(removedAllInvalidChrFromRawMaterialsDes)("sumInThisUpUsedQtyOfGoods") = _
            rawMaterialsQtyGroupByGoods(removedAllInvalidChrFromRawMaterialsDes)("sumInThisUpUsedQtyOfGoods") + upClause8InfoDic(dicKey)("inThisUpUsedQtyOfGoods")

        totalUsedQtyOfGoods = totalUsedQtyOfGoods + upClause8InfoDic(dicKey)("inThisUpUsedQtyOfGoods")
        totalUsedValueOfGoods = totalUsedValueOfGoods + upClause8InfoDic(dicKey)("inThisUpUsedValueOfGoods")

    Next dicKey

    'insert rows as goods count, note already one row exist
    Dim i As Long
    For i = 1 To rawMaterialsQtyGroupByGoods.Count - 1
        workingRange.Rows("2").EntireRow.Insert
    Next i

    Set workingRange = workingRange.Resize(rawMaterialsQtyGroupByGoods.Count)
    workingRange.Clear
    Application.Run "utility_formating_fun.rangeFormat", workingRange, "Calibri", 10, False, True, xlCenter, xlCenter, "General"


    Dim j As Long

    For j = 0 To rawMaterialsQtyGroupByGoods.Count - 1

        dicKey = rawMaterialsQtyGroupByGoods.keys()(j)

        workingRange.Range("C" & j + 1).value = j + 1
        workingRange.Range("D" & j + 1).value = rawMaterialsQtyGroupByGoods(dicKey)("nameOfGoods")
        workingRange.Range("D" & j + 1 & ":G" & j + 1).Merge
        workingRange.Range("H" & j + 1).value = Join(rawMaterialsQtyGroupByGoods(dicKey)("concatedHsCode").keys, ", ")
        workingRange.Range("H" & j + 1 & ":I" & j + 1).Merge
        workingRange.Range("J" & j + 1).value = rawMaterialsQtyGroupByGoods(dicKey)("sumInThisUpUsedQtyOfGoods")

    Next j

    workingRange.Range("K" & 10).value = totalUsedValueOfGoods / totalUsedQtyOfGoods

    Application.Run "utility_formating_fun.SetBorderThin", workingRange.Range("C1:L" & workingRange.Rows.Count)
    Application.Run "utility_formating_fun.removeBorder", workingRange.Range("K1:L" & workingRange.Rows.Count), xlInsideHorizontal
    workingRange.Range("J1:K" & workingRange.Rows.Count).Style = "Comma"

End Function