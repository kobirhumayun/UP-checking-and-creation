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