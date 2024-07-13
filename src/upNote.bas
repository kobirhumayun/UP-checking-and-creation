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

        If Left(sourceDataAsDicUpIssuingStatus(dicKey)("currencyNumberFormat"), 8) = vsCodeNotSupportedOrBengaliTxtDictionary("sourceDataAsDicUpIssuingStatusCurrencyNumberFormat") Then

            exportValue = exportValue + CDbl(Round(sourceDataAsDicUpIssuingStatus(dicKey)("LCAmount") * 1.05)) ' conversion rate would be dynamic

        Else

            exportValue = exportValue + CDbl(sourceDataAsDicUpIssuingStatus(dicKey)("LCAmount"))

        End If

        If Right(sourceDataAsDicUpIssuingStatus(dicKey)("qtyNumberFormat"), 5) = """Mtr""" Then

            exportQty = exportQty + Round(sourceDataAsDicUpIssuingStatus(dicKey)("QuantityofFabricsYdsMtr") * 1.0936132983)

        Else

            exportQty = exportQty + sourceDataAsDicUpIssuingStatus(dicKey)("QuantityofFabricsYdsMtr")

        End If

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
        Dim exportValue, exportQty As Variant
        Dim dicKey As Variant


        For j = 0 To sourceDataAsDicUpIssuingStatus.Count - 1

            dicKey = sourceDataAsDicUpIssuingStatus.keys()(j)

            exportValue = 0
            exportQty = 0

            If Left(sourceDataAsDicUpIssuingStatus(dicKey)("currencyNumberFormat"), 8) = vsCodeNotSupportedOrBengaliTxtDictionary("sourceDataAsDicUpIssuingStatusCurrencyNumberFormat") Then

                exportValue = CDbl(Round(sourceDataAsDicUpIssuingStatus(dicKey)("LCAmount") * 1.05)) ' conversion rate would be dynamic

            Else

                exportValue = CDbl(sourceDataAsDicUpIssuingStatus(dicKey)("LCAmount"))

            End If

            If Right(sourceDataAsDicUpIssuingStatus(dicKey)("qtyNumberFormat"), 5) = """Mtr""" Then

                exportQty = Round(sourceDataAsDicUpIssuingStatus(dicKey)("QuantityofFabricsYdsMtr") * 1.0936132983)

            Else

                exportQty = sourceDataAsDicUpIssuingStatus(dicKey)("QuantityofFabricsYdsMtr")

            End If

            workingRange.Range("C" & j + 1).value = j + 1
            workingRange.Range("D" & j + 1).value = Application.Run("createUp.combinLcAndAmnd", sourceDataAsDicUpIssuingStatus(dicKey))
            workingRange.Range("E" & j + 1).value = exportValue
            workingRange.Range("F" & j + 1).value = exportQty
            ' workingRange.Range("n" & j + 1 & ":z" & j + 1).Merge
        Next j


        ' workingRange.Range("b1:b" & workingRange.Rows.Count).Merge
        ' workingRange.Range("c1:m" & workingRange.Rows.Count).Merge

        ' Application.Run "utility_formating_fun.SetBorderInsideHairlineAroundThin", workingRange.Range("b1:z" & workingRange.Rows.Count)
        ' Application.Run "utility_formating_fun.setBorder", workingRange.Range("b1:z1"), xlEdgeTop, xlHairline


End Function