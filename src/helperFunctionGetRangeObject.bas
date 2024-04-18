Attribute VB_Name = "helperFunctionGetRangeObject"
Option Explicit



Private Function upClause6BuyerinformationRangeObject() As Variant
'this function give buyer information Range Object from active sheet

Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("6|", LookAt:=xlWhole).Row
bottomRow = Range("N" & Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("localB2bLcDesBengaliTxt"), LookAt:=xlPart).Row).End(xlUp).Row

Dim workingRange As Range
Set workingRange = Range("N" & topRow & ":" & "N" & bottomRow)
workingRange.Font.Color = RGB(255, 255, 255)

Set upClause6BuyerinformationRangeObject = workingRange

End Function


Private Function upClause6BuyerinformationRangeObjectFromProvidedWs(ws As Worksheet) As Variant
    'this function give buyer information Range Object from provided sheet

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find("6|", LookAt:=xlWhole).Row
    bottomRow = ws.Range("N" & ws.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("localB2bLcDesBengaliTxt"), LookAt:=xlPart).Row).End(xlUp).Row

    Dim workingRange As Range
    Set workingRange = ws.Range("A" & topRow & ":" & "AI" & bottomRow)

    Set upClause6BuyerinformationRangeObjectFromProvidedWs = workingRange

End Function


Private Function upClause7LcinformationRangeObject() As Variant
'this function give LC information Range Object from active sheet

Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("localB2bLcDesBengaliTxt"), LookAt:=xlPart).Row + 2
bottomRow = Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row - 1

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AA" & bottomRow)
workingRange.Font.Color = RGB(255, 255, 255)
Set upClause7LcinformationRangeObject = workingRange

End Function

Private Function upClause7LcinformationRangeObjectFromProvidedWs(ws As Worksheet) As Variant
    'this function give buyer lc information Range Object from provided sheet

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("localB2bLcDesBengaliTxt"), LookAt:=xlPart).Row + 1
    bottomRow = ws.Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row - 1

    Dim workingRange As Range
    Set workingRange = ws.Range("A" & topRow & ":" & "AI" & bottomRow)

    Set upClause7LcinformationRangeObjectFromProvidedWs = workingRange

End Function

Private Function upClause8BtbLcinformationRangeObject() As Variant
'this function give BTB LC information Range Object from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row + 3
bottomRow = Range("V" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AA" & bottomRow)
workingRange.Font.Color = RGB(255, 255, 255)
Set upClause8BtbLcinformationRangeObject = workingRange

End Function

Private Function upClause8BtbLcinformationRangeObjectFromProvidedWs(ws As Worksheet) As Variant
    'this function give BTB LC information Range Object from provided sheet
    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row + 3
    bottomRow = ws.Range("V" & topRow).End(xlDown).Row

    Dim workingRange As Range
    Set workingRange = ws.Range("B" & topRow & ":" & "AA" & bottomRow)
    workingRange.Font.Color = RGB(0, 255, 0)
    Set upClause8BtbLcinformationRangeObjectFromProvidedWs = workingRange

End Function

Private Function upClause8BtbLcinformationRangeObjectPreviousFormatFromProvidedWs(ws As Worksheet) As Variant
    'this function give BTB LC information Range Object from provided sheet
    Dim topRow, bottomRow As Variant

    topRow = ws.Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row + 2
    bottomRow = ws.Range("S" & topRow).End(xlDown).Row

    Dim workingRange As Range
    Set workingRange = ws.Range("B" & topRow & ":" & "T" & bottomRow)
    workingRange.Font.Color = RGB(0, 255, 0)
    Set upClause8BtbLcinformationRangeObjectPreviousFormatFromProvidedWs = workingRange

End Function
 

Private Function upClause9StockinformationRangeObject() As Variant
'this function give stock information Range Object from active sheet

Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("infoAboutStockBengaliTxt"), LookAt:=xlPart).Row + 3
bottomRow = Range("T" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AC" & bottomRow)
workingRange.Font.Color = RGB(255, 255, 255)
Set upClause9StockinformationRangeObject = workingRange

End Function

Private Function upClause11UdExpIpinformationRangeObject() As Variant
'this function give UD/EXP/IP information Range Object from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("11|", LookAt:=xlPart).Row + 3
bottomRow = Range("Z" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AA" & bottomRow)
workingRange.Font.Color = RGB(255, 255, 255)
Set upClause11UdExpIpinformationRangeObject = workingRange

End Function

Private Function upClause12AYarnConsumptioninformationRangeObject() As Variant
'this function give yarn consumption information Range Object from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("12| (K)", LookAt:=xlPart).Row + 2
bottomRow = Range("Z" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AA" & bottomRow)
workingRange.Font.Color = RGB(255, 255, 255)
Set upClause12AYarnConsumptioninformationRangeObject = workingRange

End Function

Private Function upClause12BChemicalDyesConsumptioninformationRangeObject() As Variant
'this function give chemical & dyes consumption information Range Object from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("12| (L)", LookAt:=xlPart).Row + 2
bottomRow = Range("X" & topRow + 1).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "Y" & bottomRow)
workingRange.Font.Color = RGB(255, 255, 255)
Set upClause12BChemicalDyesConsumptioninformationRangeObject = workingRange

End Function

Private Function upClause13UseRawMaterialsinformationRangeObject() As Variant
'this function give used raw materials information Range Object from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("13|", LookAt:=xlPart).Row + 2
bottomRow = Range("R" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "R" & bottomRow)
workingRange.Font.Color = RGB(255, 255, 255)
Set upClause13UseRawMaterialsinformationRangeObject = workingRange


End Function



Private Function sourceDataImportPerformanceRangeObject(fileName As String, worksheetTabName As String, openFile As Boolean, closeFile As Boolean) As Variant ' provide source file name & worksheetTabName
    'this function give source data from Import Performance
        
        
        
        If openFile Then
            Application.Run "utilityFunction.openFile", fileName ' provide filename
        End If
        
        ActiveWorkbook.Worksheets(worksheetTabName).Activate
        ActiveSheet.AutoFilterMode = False
        
        Dim topRow, bottomRow As Variant

        topRow = 5
        bottomRow = Range("C5").End(xlDown).Row

        Dim workingRange As Range
        Set workingRange = Range("A" & topRow & ":" & "N" & bottomRow)
        
        
        If closeFile Then
           Application.Run "utilityFunction.closeFile", fileName ' provide filename
        End If


        
        Set sourceDataImportPerformanceRangeObject = workingRange
    
End Function
    
