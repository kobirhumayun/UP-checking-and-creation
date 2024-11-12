Attribute VB_Name = "previousFormatRelatedFun"
Option Explicit

Private Function upClause8InformationFromProvidedWsPrevFormat(ws As Worksheet) As Object
    'this function give source data as dictionary from UP clause8

    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row + 3
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

        propertiesValArr(1) = temp(i, 2)
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

        tempMuOrBillKey = Application.Run("general_utility_functions.dictKeyGeneratorWithLcMushakOrBillOfEntryQtyAndValue", temp(i, 2), temp(i, 7), temp(i, 16), temp(i, 17))

        upClause8Dic.Add tempMuOrBillKey, tempMushakOrBillOfEntryDic

    Next i


    Set upClause8InformationFromProvidedWsPrevFormat = upClause8Dic

End Function

Private Function upClause6BuyerinformationRangeObjectFromProvidedWsPrevFormat(ws As Worksheet) As Variant
    'this function give buyer information Range Object from provided sheet

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find("6|", LookAt:=xlWhole).Row
    bottomRow = ws.Range("N" & ws.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("localB2bLcDesBengaliTxtPrevFormat"), LookAt:=xlPart).Row).End(xlUp).Row

    Dim workingRange As Range
    Set workingRange = ws.Range("A" & topRow & ":" & "AI" & bottomRow)

    Set upClause6BuyerinformationRangeObjectFromProvidedWsPrevFormat = workingRange

End Function

Private Function upClause7LcinformationRangeObjectFromProvidedWsPrevFormat(ws As Worksheet) As Variant
    'this function give buyer lc information Range Object from provided sheet

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("localB2bLcDesBengaliTxtPrevFormat"), LookAt:=xlPart).Row + 1
    bottomRow = ws.Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row - 1

    Dim workingRange As Range
    Set workingRange = ws.Range("A" & topRow & ":" & "AI" & bottomRow)

    Set upClause7LcinformationRangeObjectFromProvidedWsPrevFormat = workingRange

End Function

Private Function upClause12AYarnConsumptioninformationRangeObjectFromProvidedWsPrevFormat(ws As Worksheet) As Variant
    'give yarn consumption information Range Object from provided sheet
    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find("12| (K)", LookAt:=xlPart).Row + 2
    bottomRow = ws.Range("Z" & topRow).End(xlDown).Row

    Dim workingRange As Range
    Set workingRange = ws.Range("A" & topRow & ":" & "AA" & bottomRow)

    workingRange.Font.Color = RGB(255, 255, 255)
    
    Set upClause12AYarnConsumptioninformationRangeObjectFromProvidedWsPrevFormat = workingRange

End Function

Private Function upClause15RangeObjectFromProvidedWsPrevFormat(ws As Worksheet) As Variant
    'give used raw materials information Range Object from provided sheet
    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find("15|", LookAt:=xlPart).Row
    bottomRow = topRow + 3

    Dim workingRange As Range
    Set workingRange = ws.Range("A" & topRow & ":" & "AA" & bottomRow)

    workingRange.Font.Color = RGB(255, 255, 255)

    Set upClause15RangeObjectFromProvidedWsPrevFormat = workingRange

End Function