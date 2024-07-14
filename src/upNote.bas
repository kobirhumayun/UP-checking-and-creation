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

        Dim j, l, m As Long
        Dim dicKey As Variant
        Dim innerDicKey As Variant

        Dim tempWidthStr As Object
        Dim tempWeightStr As Object
        Dim temp As Variant


        For j = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

            dicKey = sourceDataAsDicUpIssuingStatus.keys()(j)

            workingRange.Range("C" & j + 1).value = Application.Run("createUp.combinUdIpExpAndDt", sourceDataAsDicUpIssuingStatus(dicKey))
            workingRange.Range("C" & j + 1 & ":G" & j + 1).Merge

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