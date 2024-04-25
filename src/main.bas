Attribute VB_Name = "main"
Option Explicit

'Public finalRawMaterialsQtyDicAsGroup As Object
'Public nnws As Worksheet

Sub checkUpWithIpExpUdMlc()

'Dim t As Single
't = Timer


Application.ScreenUpdating = False


On Error GoTo ErrorMsg

Dim upWorkBook As Workbook
Dim upWorksheet As Worksheet
Dim consumptionWorksheet As Worksheet

Set upWorkBook = ActiveWorkbook
Set upWorksheet = upWorkBook.Worksheets(2)
Set consumptionWorksheet = upWorkBook.Worksheets("Consumption")

'take UP no.
Dim upNo As Variant
upNo = Application.Run("helperFunctionGetData.upNo")


'take buyer information from UP clause 6
Dim upClause6Buyerinformation As Variant
upClause6Buyerinformation = Application.Run("helperFunctionGetData.upClause6BuyerInformation")
'take buyer information RangeObject from UP clause 6
Dim upClause6BuyerinformationRangeObject As Variant
Set upClause6BuyerinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause6BuyerInformationRangeObject")


'take LC information from UP clause 7
Dim upClause7Lcinformation As Variant
upClause7Lcinformation = Application.Run("helperFunctionGetData.upClause7LcInformation")
'take LC information RangeObject from UP clause 7
Dim upClause7LcinformationRangeObject As Variant
Set upClause7LcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause7LcInformationRangeObject")


'take BTB LC information from UP clause 8
Dim upClause8BtbLcinformation As Variant
upClause8BtbLcinformation = Application.Run("helperFunctionGetData.upClause8BtbLcInformation")
'take BTB LC information RangeObject from UP clause 8
Dim upClause8BtbLcinformationRangeObject As Variant
Set upClause8BtbLcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause8BtbLcInformationRangeObject")


'take stock information from UP clause 9
Dim upClause9Stockinformation As Variant
upClause9Stockinformation = Application.Run("helperFunctionGetData.upClause9StockInformation")
'take stock information RangeObject from UP clause 9
Dim upClause9StockinformationRangeObject As Variant
Set upClause9StockinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause9StockInformationRangeObject")


'take UD/EXP/IP information from UP clause 11
Dim upClause11UdExpIpinformation As Variant
upClause11UdExpIpinformation = Application.Run("helperFunctionGetData.upClause11UdExpIpInformation")
'take UD/EXP/IP information RangeObject from UP clause 11
Dim upClause11UdExpIpinformationRangeObject As Variant
Set upClause11UdExpIpinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause11UdExpIpInformationRangeObject")


'take yarn consumption information from UP clause 12(a)
Dim upClause12AYarnConsumptioninformation As Variant
upClause12AYarnConsumptioninformation = Application.Run("helperFunctionGetData.upClause12AYarnConsumptionInformation")
'take yarn consumption information RangeObject from UP clause 12(a)
Dim upClause12AYarnConsumptioninformationRangeObject As Variant
Set upClause12AYarnConsumptioninformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause12AYarnConsumptionInformationRangeObject")


'take chemical & dyes consumption information from UP clause 12(b)
Dim upClause12BChemicalDyesConsumptioninformation As Variant
upClause12BChemicalDyesConsumptioninformation = Application.Run("helperFunctionGetData.upClause12BChemicalDyesConsumptionInformation")
'take chemical & dyes consumption information RangeObject from UP clause 12(b)
Dim upClause12BChemicalDyesConsumptioninformationRangeObject As Variant
Set upClause12BChemicalDyesConsumptioninformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause12BChemicalDyesConsumptionInformationRangeObject")


'take raw materials information from UP clause 13
Dim upClause13UseRawMaterialsinformation As Variant
upClause13UseRawMaterialsinformation = Application.Run("helperFunctionGetData.upClause13UseRawMaterialsInformation")
'take raw materials information RangeObject from UP clause 13
Dim upClause13UseRawMaterialsinformationRangeObject As Variant
Set upClause13UseRawMaterialsinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause13UseRawMaterialsInformationRangeObject")


'take yarn consumption information from UP consumption sheet
Dim upYarnConsumptionInformation As Variant
upYarnConsumptionInformation = Application.Run("helperFunctionGetData.upYarnConsumptionInformation")


'take yarn consumption info from "Consumption" sheet
Dim yarnConsumptionInfoDic As Variant
Set yarnConsumptionInfoDic = Application.Run("afterConsumption.upYarnConsumptionInformationFromProvidedWs", consumptionWorksheet)


'take source data from UP Issuing Status
Dim sourceDataUpIssuingStatus As Variant
sourceDataUpIssuingStatus = Application.Run("helperFunctionGetData.SourceDataUPIssuingStatus", upNo, "UP Issuing Status for the Period # 01-03-2024 to 28-02-2025.xlsx", "UP Issuing Status # 2024-2025")

'take source data as dictionary from UP Issuing Status
Dim sourceDataAsDicUpIssuingStatus As Variant
Set sourceDataAsDicUpIssuingStatus = Application.Run("helperFunctionGetData.sourceDataAsDicUpIssuingStatus", upNo, "UP Issuing Status for the Period # 01-03-2024 to 28-02-2025.xlsx", "UP Issuing Status # 2024-2025")


'take source data from Import Performance Yarn Import
Dim sourceDataImportPerformanceYarnImport As Variant
sourceDataImportPerformanceYarnImport = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2024-2025.xlsx", "Yarn (Import)", True, False)


'take source data from Import Performance Yarn Local
Dim sourceDataImportPerformanceYarnLocal As Variant
sourceDataImportPerformanceYarnLocal = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2024-2025.xlsx", "Yarn (Local)", False, False)


'take source data from Import Performance Dyes
Dim sourceDataImportPerformanceDyes As Variant
sourceDataImportPerformanceDyes = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2024-2025.xlsx", "Dyes", False, False)


'take source data from Import Performance Chemicals Import
Dim sourceDataImportPerformanceChemicalsImport As Variant
sourceDataImportPerformanceChemicalsImport = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2024-2025.xlsx", "Chemicals (Import)", False, False)


'take source data from Import Performance Chemicals Local
Dim sourceDataImportPerformanceChemicalsLocal As Variant
sourceDataImportPerformanceChemicalsLocal = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2024-2025.xlsx", "Chemicals (Local)", False, False)


'take source data from Import Performance Stretch Wrapping Film
Dim sourceDataImportPerformanceStretchWrappingFilm As Variant
sourceDataImportPerformanceStretchWrappingFilm = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2024-2025.xlsx", "St.Wrap.Film (Import)", False, False)


'take source data from Import Performance Total Summary
Dim sourceDataImportPerformanceTotalSummary As Variant
sourceDataImportPerformanceTotalSummary = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2024-2025.xlsx", "Summary of Grand Total", False, True)



'take source data from previous UP clause 8
Dim sourceDataPreviousUpClause8 As Variant
sourceDataPreviousUpClause8 = Application.Run("helperFunctionGetData.sourceDataPreviousUp", upNo, 8, True, False)



'take source data from previous UP clause 9
Dim sourceDataPreviousUpClause9 As Variant
sourceDataPreviousUpClause9 = Application.Run("helperFunctionGetData.sourceDataPreviousUp", upNo, 9, False, True)






'#####compare section start from here#####

Dim upClause6And7CompareWithSourceResultArr As Variant
upClause6And7CompareWithSourceResultArr = Application.Run("helperFunctionCompareData.upClause6And7CompareWithSource", upClause6BuyerinformationRangeObject, upClause7LcinformationRangeObject, sourceDataUpIssuingStatus)


Dim upClause8CompareWithSourceResultArr As Variant
upClause8CompareWithSourceResultArr = Application.Run("helperFunctionCompareData.upClause8CompareWithSource", upClause8BtbLcinformationRangeObject, sourceDataUpIssuingStatus, sourceDataImportPerformanceYarnImport, sourceDataImportPerformanceYarnLocal, sourceDataImportPerformanceDyes, sourceDataImportPerformanceChemicalsImport, sourceDataImportPerformanceChemicalsLocal, sourceDataImportPerformanceStretchWrappingFilm, sourceDataPreviousUpClause8, yarnConsumptionInfoDic, sourceDataAsDicUpIssuingStatus)


Dim upClause9CompareWithSourceResultArr As Variant
upClause9CompareWithSourceResultArr = Application.Run("helperFunctionCompareData.upClause9CompareWithSource", upClause9StockinformationRangeObject, upClause8CompareWithSourceResultArr, upYarnConsumptionInformation, sourceDataPreviousUpClause9, sourceDataImportPerformanceTotalSummary)


Dim upClause11CompareWithSourceResultArr As Variant
upClause11CompareWithSourceResultArr = Application.Run("helperFunctionCompareData.upClause11CompareWithSource", upClause11UdExpIpinformationRangeObject, upClause6Buyerinformation, upClause7Lcinformation, upClause6And7CompareWithSourceResultArr, sourceDataUpIssuingStatus)


Dim upClause12aCompareWithSourceResultArr As Variant
upClause12aCompareWithSourceResultArr = Application.Run("helperFunctionCompareData.upClause12aCompareWithSource", upClause12AYarnConsumptioninformationRangeObject, upClause6Buyerinformation, upClause7Lcinformation, upYarnConsumptionInformation)


Dim upClause12bCompareWithSourceResultArr As Variant
upClause12bCompareWithSourceResultArr = Application.Run("helperFunctionCompareData.upClause12bCompareWithSource", upClause12BChemicalDyesConsumptioninformationRangeObject, upClause6Buyerinformation, upYarnConsumptionInformation, upClause7Lcinformation)


Dim upClause13CompareWithSourceResultArr As Variant
upClause13CompareWithSourceResultArr = Application.Run("helperFunctionCompareData.upClause13CompareWithSource", upClause13UseRawMaterialsinformationRangeObject, upClause8CompareWithSourceResultArr)




'Result put to result sheet  start
Sheets.Add After:=Sheets(ActiveSheet.Name)


Dim upClause6And7CompareWithSourceResultPutRange As Range
Set upClause6And7CompareWithSourceResultPutRange = ActiveSheet.Range("a" & Cells.SpecialCells(xlCellTypeLastCell).Row).Resize(UBound(upClause6And7CompareWithSourceResultArr, 1), UBound(upClause6And7CompareWithSourceResultArr, 2))
upClause6And7CompareWithSourceResultPutRange = upClause6And7CompareWithSourceResultArr

Dim upClause8CompareWithSourceResultPutRange As Range
Set upClause8CompareWithSourceResultPutRange = ActiveSheet.Range("a" & Cells.SpecialCells(xlCellTypeLastCell).Row + 2).Resize(UBound(upClause8CompareWithSourceResultArr, 1), UBound(upClause8CompareWithSourceResultArr, 2))
upClause8CompareWithSourceResultPutRange = upClause8CompareWithSourceResultArr


Dim upClause9CompareWithSourceResultPutRange As Range
Set upClause9CompareWithSourceResultPutRange = ActiveSheet.Range("a" & Cells.SpecialCells(xlCellTypeLastCell).Row + 2).Resize(UBound(upClause9CompareWithSourceResultArr, 1), UBound(upClause9CompareWithSourceResultArr, 2))
upClause9CompareWithSourceResultPutRange = upClause9CompareWithSourceResultArr



Dim upClause11CompareWithSourceResultPutRange As Range
Set upClause11CompareWithSourceResultPutRange = ActiveSheet.Range("a" & Cells.SpecialCells(xlCellTypeLastCell).Row + 2).Resize(UBound(upClause11CompareWithSourceResultArr, 1), UBound(upClause11CompareWithSourceResultArr, 2))
upClause11CompareWithSourceResultPutRange = upClause11CompareWithSourceResultArr



Dim upClause12aCompareWithSourceResultPutRange As Range
Set upClause12aCompareWithSourceResultPutRange = ActiveSheet.Range("a" & Cells.SpecialCells(xlCellTypeLastCell).Row + 2).Resize(UBound(upClause12aCompareWithSourceResultArr, 1), UBound(upClause12aCompareWithSourceResultArr, 2))
upClause12aCompareWithSourceResultPutRange = upClause12aCompareWithSourceResultArr



Dim upClause12bCompareWithSourceResultPutRange As Range
Set upClause12bCompareWithSourceResultPutRange = ActiveSheet.Range("a" & Cells.SpecialCells(xlCellTypeLastCell).Row + 2).Resize(UBound(upClause12bCompareWithSourceResultArr, 1), UBound(upClause12bCompareWithSourceResultArr, 2))
upClause12bCompareWithSourceResultPutRange = upClause12bCompareWithSourceResultArr



Dim upClause13CompareWithSourceResultPutRange As Range
Set upClause13CompareWithSourceResultPutRange = ActiveSheet.Range("a" & Cells.SpecialCells(xlCellTypeLastCell).Row + 2).Resize(UBound(upClause13CompareWithSourceResultArr, 1), UBound(upClause13CompareWithSourceResultArr, 2))
upClause13CompareWithSourceResultPutRange = upClause13CompareWithSourceResultArr




Application.Run "utilityFunction.resultSheetFormating", upClause6And7CompareWithSourceResultPutRange, RGB(2, 34, 443), RGB(255, 255, 255), RGB(0, 0, 0)

Application.Run "utilityFunction.resultSheetFormating", upClause8CompareWithSourceResultPutRange, RGB(1, 3, 34), RGB(255, 255, 255), RGB(105, 0, 0)

Application.Run "utilityFunction.resultSheetFormating", upClause9CompareWithSourceResultPutRange, RGB(2, 34, 443), RGB(255, 255, 255), RGB(0, 0, 0)

Application.Run "utilityFunction.resultSheetFormating", upClause11CompareWithSourceResultPutRange, RGB(1, 3, 34), RGB(255, 255, 255), RGB(105, 0, 0)

Application.Run "utilityFunction.resultSheetFormating", upClause12aCompareWithSourceResultPutRange, RGB(2, 34, 443), RGB(255, 255, 255), RGB(0, 0, 0)

Application.Run "utilityFunction.resultSheetFormating", upClause12bCompareWithSourceResultPutRange, RGB(1, 3, 34), RGB(255, 255, 255), RGB(105, 0, 0)

Application.Run "utilityFunction.resultSheetFormating", upClause13CompareWithSourceResultPutRange, RGB(2, 34, 443), RGB(255, 255, 255), RGB(0, 0, 0)



'Result put to result sheet end



''Debug.Print Timer - t, Timer

Application.ScreenUpdating = True

Exit Sub

ErrorMsg:
MsgBox "Operation not completed, may you get the wrong result."

End Sub






'
'Sub onlyForReport()
'
''Dim t As Single
''t = Timer
'
'
'Application.ScreenUpdating = False
'
'
'On Error GoTo ErrorMsg
'
'
'
'Dim i As Integer
'
'For i = 1 To Sheets.Count
'
'
'Workbooks.Open fileName:=ActiveWorkbook.path & Application.PathSeparator & "Audit Period # 2024-2025 All UP File Merged.xlsx", ReadOnly:=False
'
'
'Worksheets(i).Select
'
''take UP no.
'Dim upNo As Variant
'upNo = Application.Run("helperFunctionGetData.upNo")
'
'
''take BTB LC information from UP clause 8
'Dim upClause8BtbLcinformation As Variant
'upClause8BtbLcinformation = Application.Run("helperFunctionGetData.upClause8BtbLcInformation")
''take BTB LC information RangeObject from UP clause 8
'Dim upClause8BtbLcinformationRangeObject As Variant
'Set upClause8BtbLcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause8BtbLcInformationRangeObject")
'
'Application.Run "utilityFunction.upClause8Report", upClause8BtbLcinformationRangeObject, upNo ' only for audit report
'
'Next
'
'Application.ScreenUpdating = True
'
'Exit Sub
'
'ErrorMsg:
'MsgBox "Operation not completed, may you get the wrong result."
'
'End Sub

'
'
'Sub putDyesCatagoryAsImport()
'
'
'
'Application.ScreenUpdating = False
'
'
'On Error GoTo ErrorMsg
'
''take source data from Import Performance Dyes
'Dim sourceDataImportPerformanceDyes As Variant
'sourceDataImportPerformanceDyes = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2024-2025.xlsx", "Dyes", True, True)
'
'Dim i As Integer
'
'For i = 1 To Sheets.Count
'
'
'Worksheets(i).Select
''take UP no.
'Dim upNo As Variant
'upNo = Application.Run("helperFunctionGetData.upNo")
'
'
''take BTB LC information from UP clause 8
'Dim upClause8BtbLcinformation As Variant
'upClause8BtbLcinformation = Application.Run("helperFunctionGetData.upClause8BtbLcInformation")
''take BTB LC information RangeObject from UP clause 8
'Dim upClause8BtbLcinformationRangeObject As Variant
'Set upClause8BtbLcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause8BtbLcInformationRangeObject")
'
'
'
'
'Application.Run "utilityFunction.putDyesCatagoryAsImportPerformance", upClause8BtbLcinformationRangeObject, upNo, sourceDataImportPerformanceDyes ' only for put dyes catagory as import performance
'
'
'
'    Next
'
'Application.ScreenUpdating = True
'
'Exit Sub
'
'ErrorMsg:
'MsgBox "Operation not completed, may you get the wrong result."
'
'End Sub
'
'
'
'Sub putAllCatagoryAsImport()
'
'
'
'    Application.ScreenUpdating = False
'
'
'    On Error GoTo ErrorMsg
'
'    'take source data from Import Performance Dyes
'    Dim sourceDataImportPerformanceAll As Variant
'    sourceDataImportPerformanceAll = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2024-2025.xlsx", "All", True, True)
'
'    Dim i As Integer
'
'    For i = 1 To Sheets.Count
'
'
'    Worksheets(i).Select
'    'take UP no.
'    Dim upNo As Variant
'    upNo = Application.Run("helperFunctionGetData.upNo")
'
'
'    'take BTB LC information from UP clause 8
'    Dim upClause8BtbLcinformation As Variant
'    upClause8BtbLcinformation = Application.Run("helperFunctionGetData.upClause8BtbLcInformation")
'
'
'    'take BTB LC information RangeObject from UP clause 8
'    Dim upClause8BtbLcinformationRangeObject As Variant
'    Set upClause8BtbLcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause8BtbLcInformationRangeObject")
'
'
'
'
'    Application.Run "utilityFunction.putAllCatagoryAsImportPerformance", upClause8BtbLcinformationRangeObject, upNo, sourceDataImportPerformanceAll ' only for put dyes catagory as import performance
'
'
'
'        Next
'
'    Application.ScreenUpdating = True
'
'    Exit Sub
'
'ErrorMsg:
'    MsgBox "Operation not completed, may you get the wrong result."
'
'    End Sub
    
Sub totalPeriodRawMaterialsUsedReport()

    Application.ScreenUpdating = False
    
    On Error GoTo ErrorMsg
    
    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    Dim totalClassifiedDict As Object
    Set totalClassifiedDict = CreateObject("Scripting.Dictionary") ' raw materials classification on report
    
    Set totalClassifiedDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", totalClassifiedDict, 0, _
    Array("Foreign Yarn (KGS)", "Local Yarn (KGS)", "Total Yarn", "Wetting Agent", "Modified Starch", "Caustic Soda", "Sulphuric Acid", _
    "Reducing Agent", "Softener", "Binder", "Sequestering Agent", "Sodium Hydro Sulphate", "Wax", "Acetic Acid", "PVA", "Desizing Agent /  Enzyme", _
    "Fixing Agent", "Dispersing Agent", "Alum+ Cataionic", "Water Decoloring Agent.", "Hydrogen Peroxide", "Stabilizing Agent", "Detergent", _
    "Resign", "Total Chemicals", "Vat Dyes  (Liquid)", "Vat Dyes (Indigo Granular)", "Sulphur Dyes (Liquid)", "Sulphur Dyes (Sulphur Granular)", "Stretch Wrapping Film"))
    
    Dim totalClassifiedDictKeys As Variant
    totalClassifiedDictKeys = totalClassifiedDict.keys
    
    Dim useGroupDict As Object
    Set useGroupDict = CreateObject("Scripting.Dictionary") ' UP raw materials group against report classification alternatively UP raw materials assign report id
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(0), _
    Array("Foreign Yarn"))


    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(1), _
    Array("Local Yarn"))


    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(2), _
    Array("Total Yarn"))

    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(3), _
    Array("Wetting Agent", "Mercerizing Agent (Wetting Agent)"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(4), _
    Array("Modified Starch", "Modified Starches", "Finishing Agent", "Finishing Agent(Modified Starch/ Sizing Agent)"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(5), _
    Array("Caustic Soda"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(6), _
    Array("Sulphuric Acid"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(7), _
    Array("Reducing Agent"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(8), _
    Array("Softener", "Softening Agent (Softener)"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(9), _
    Array("Binder"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(10), _
    Array("Sequestering Agent"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(11), _
    Array("Sodium Hydro Sulphate", "Sodium Hydro Sulphite"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(12), _
    Array("Wax", "Waxes"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(13), _
    Array("Acetic Acid", "Acetic Acid/Green Acid"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(14), _
    Array("PVA"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(15), _
    Array("Desizing Agent /  Enzyme", "Desizing Agent", "Enzyme"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(16), _
    Array("Fixing Agent"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(17), _
    Array("Dispersing Agent"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(18), _
    Array("Alum+ Cataionic"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(19), _
    Array("Water Decoloring Agent."))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(20), _
    Array("Hydrogen Peroxide"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(21), _
    Array("Stabilizing Agent", "Stabilizing Agent (Estabilizador FE)", "Finishing Agent (Estabilizador FE)"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(22), _
    Array("Detergent"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(23), _
    Array("Resign"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(24), _
    Array("Total Chemicals"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(25), _
    Array("Vat Dyes"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(26), _
    Array("Vat Dyes (Indigo Granular)"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(27), _
    Array("Sulphur Dyes"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(28), _
    Array("Sulphur Dyes (Sulphur Granular)"))
    
    
    Set useGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", useGroupDict, totalClassifiedDictKeys(29), _
    Array("Stretch Wrapping Film"))
    
    Dim importPerformanceFilePath As String
    importPerformanceFilePath = ActiveWorkbook.path & Application.PathSeparator & "Import Performance Statement of PDL-2024-2025.xlsx" ' file name will be change after change period
    
    Dim impBillAndMushakDb As Object
    Set impBillAndMushakDb = Application.Run("utilityFunction.CombinedAllSheetsMushakOrBillOfEntryDbDict", importPerformanceFilePath)
    
    Dim inputTxt As Variant
    inputTxt = InputBox("Please enter a Text ""UP"" or ""Import"" ", "Description Take From ""UP"" Or ""Import Performance""", "UP")

    ' merged file iterate start
    
    Dim allUpFinalResultDict As Object
    Set allUpFinalResultDict = CreateObject("Scripting.Dictionary") 'final result DB
    
    Dim rawMaterialsNotClassifiedDict As Object
    Set rawMaterialsNotClassifiedDict = CreateObject("Scripting.Dictionary") ' not classified raw materials container of final result DB
    
    allUpFinalResultDict.Add "rawMaterialsNotClassifiedDict", rawMaterialsNotClassifiedDict
    
    Dim mergedUpWb As Workbook
    Dim mergedUpWs As Worksheet
    Set mergedUpWb = ActiveWorkbook

    
    For Each mergedUpWs In mergedUpWb.Worksheets


        Dim isColumn8CurrentFormat As Boolean

        isColumn8CurrentFormat = Application.Run("utilityFunction.DoesStringExistInWorksheets", vsCodeNotSupportedOrBengaliTxtDictionary("totalUsedRawMetarialsBengaliTxt"), mergedUpWs)

        'take UP no.
        Dim upNo As Variant
        upNo = Application.Run("helperFunctionGetData.upNoFromProvidedWs", mergedUpWs)


        Dim upClause8BtbLcinformationRangeObject As Variant

        If isColumn8CurrentFormat Then


            'take BTB LC information RangeObject from UP clause 8
            Set upClause8BtbLcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause8BtbLcinformationRangeObjectFromProvidedWs", mergedUpWs)

        Else


            'take BTB LC information RangeObject from UP clause 8
            Set upClause8BtbLcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause8BtbLcinformationRangeObjectPreviousFormatFromProvidedWs", mergedUpWs)

        End If

        Set allUpFinalResultDict = Application.Run("utilityFunction.addAllUpToFinalDbDictionary", upClause8BtbLcinformationRangeObject, upNo, isColumn8CurrentFormat, allUpFinalResultDict, totalClassifiedDictKeys, useGroupDict, impBillAndMushakDb, inputTxt)       'UP clause 8 add to final db dictionary
        
    Next mergedUpWs

    Dim rawMaterialFilePath As String
    rawMaterialFilePath = ActiveWorkbook.path & Application.PathSeparator & "Raw-Materials Used (Yarn, Dyes & Chemicals)_Calculation Sheet_2024-2025.xlsx" ' file name will be change after change period
    
    Dim rawMaterialWb As Workbook
    Dim rawMaterialWs As Worksheet
    Set rawMaterialWb = Workbooks.Open(rawMaterialFilePath)
    Set rawMaterialWs = rawMaterialWb.Worksheets(1)
    
    Dim tempRange As Range
    Dim topRow, bottomRow As Long
    topRow = 5
    bottomRow = Cells(topRow, 1).End(xlDown).Row
    Dim i As Long
    For i = topRow To bottomRow
        Set tempRange = rawMaterialWs.Cells(i, 1)
        
        If allUpFinalResultDict.Exists(tempRange.value) Then
            Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", tempRange.Offset(0, 1), allUpFinalResultDict(tempRange.value), 0, 1, 0
        End If
        
    Next i
    Application.ScreenUpdating = True
    
    Dim notFoundItems As Variant
    notFoundItems = allUpFinalResultDict("rawMaterialsNotClassifiedDict").items ' only for watching purpose
    
    Exit Sub

ErrorMsg:
    MsgBox "Operation not completed, may you get the wrong result."
  
   

End Sub


Sub totalPeriodBillOfEntryOrMushkUsedCalculationAndPutToImportPerformance()

  'Dim t As Single
  't = Timer
  
  
  Application.ScreenUpdating = False
  
  
  On Error GoTo ErrorMsg

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")
  
  Dim mergedUpWb As Workbook
  Dim mergedUpWs As Worksheet
  Set mergedUpWb = ActiveWorkbook
  
  ' create file start
  Dim filePath As String
  filePath = ActiveWorkbook.path & Application.PathSeparator & "All_UP_Clause8_Merged.xlsx" 'Replace with desired file path and name
  
    If Dir(filePath) <> "" Then 'Check if file not exists
    'File does exist, so delete the file
    
    Kill filePath
    
    End If
  
  If Dir(filePath) = "" Then 'Check if file exists
      'File does not exist, so create a new workbook and save it to the specified file path
      Dim newWorkbook As Workbook
      Set newWorkbook = Workbooks.Add
      newWorkbook.SaveAs filePath
      newWorkbook.Close

  End If
  ' create file end
  
  Dim helperFileWb As Workbook
  Dim helperFileWs As Worksheet
  Set helperFileWb = Workbooks.Open(filePath)
  Set helperFileWs = helperFileWb.Worksheets(1)
  
  For Each mergedUpWs In mergedUpWb.Worksheets
  
  
      Dim isColumn8CurrentFormat As Boolean
  
      isColumn8CurrentFormat = Application.Run("utilityFunction.DoesStringExistInWorksheets", vsCodeNotSupportedOrBengaliTxtDictionary("totalUsedRawMetarialsBengaliTxt"), mergedUpWs)
  
      'take UP no.
      Dim upNo As Variant
      upNo = Application.Run("helperFunctionGetData.upNoFromProvidedWs", mergedUpWs)
  

      Dim upClause8BtbLcinformationRangeObject As Variant
  
      If isColumn8CurrentFormat Then
  

          'take BTB LC information RangeObject from UP clause 8
          Set upClause8BtbLcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause8BtbLcinformationRangeObjectFromProvidedWs", mergedUpWs)
  
      Else
  

          'take BTB LC information RangeObject from UP clause 8
          Set upClause8BtbLcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause8BtbLcinformationRangeObjectPreviousFormatFromProvidedWs", mergedUpWs)
  
      End If
  
      Application.Run "utilityFunction.putAllUpToHelperFile", upClause8BtbLcinformationRangeObject, upNo, helperFileWs, isColumn8CurrentFormat    'UP clause 8 put to helper file
  
  Next mergedUpWs
  

    Dim afterMergedClause8OfAllUp As Variant
    afterMergedClause8OfAllUp = helperFileWs.Range("a2:aa" & helperFileWs.Cells.SpecialCells(xlCellTypeLastCell).Row).value
    
    helperFileWb.Close SaveChanges:=True
    
    Dim importPerformanceFileWb As Workbook
    Dim importPerformanceFilePath As String
    importPerformanceFilePath = ActiveWorkbook.path & Application.PathSeparator & "Import Performance Statement of PDL-2024-2025 for Bond Audit.xlsx" ' file name will be change after change period
    Set importPerformanceFileWb = Workbooks.Open(importPerformanceFilePath)
        
    Dim importPerformanceFileYarnImportWs As Worksheet
    Set importPerformanceFileYarnImportWs = importPerformanceFileWb.Worksheets("Yarn (Import)")

    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFile", importPerformanceFileYarnImportWs, 3, 7, 8, 9, 10, 28, afterMergedClause8OfAllUp     'Used Qty & Value put to import performance file
             
    Dim importPerformanceFileYarnLocalWs As Worksheet
    Set importPerformanceFileYarnLocalWs = importPerformanceFileWb.Worksheets("Yarn (Local)")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFile", importPerformanceFileYarnLocalWs, 3, 7, 8, 9, 10, 28, afterMergedClause8OfAllUp     'Used Qty & Value put to import performance file
        
    Dim importPerformanceFileDyesWs As Worksheet
    Set importPerformanceFileDyesWs = importPerformanceFileWb.Worksheets("Dyes")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFile", importPerformanceFileDyesWs, 3, 7, 8, 9, 10, 28, afterMergedClause8OfAllUp     'Used Qty & Value put to import performance file
        
    Dim importPerformanceFileChemialsImportWs As Worksheet
    Set importPerformanceFileChemialsImportWs = importPerformanceFileWb.Worksheets("Chemicals (Import)")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFile", importPerformanceFileChemialsImportWs, 3, 8, 9, 10, 11, 28, afterMergedClause8OfAllUp     'Used Qty & Value put to import performance file
        
    Dim importPerformanceFileChemialsLocalWs As Worksheet
    Set importPerformanceFileChemialsLocalWs = importPerformanceFileWb.Worksheets("Chemicals (Local)")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFile", importPerformanceFileChemialsLocalWs, 3, 8, 9, 10, 11, 28, afterMergedClause8OfAllUp     'Used Qty & Value put to import performance file
        
    Dim importPerformanceFileWrappingFilmWs As Worksheet
    Set importPerformanceFileWrappingFilmWs = importPerformanceFileWb.Worksheets("St.Wrap.Film (Import)")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFile", importPerformanceFileWrappingFilmWs, 3, 8, 9, 10, 11, 28, afterMergedClause8OfAllUp     'Used Qty & Value put to import performance file

   
  Application.ScreenUpdating = True
  
  Exit Sub
  
ErrorMsg:
  MsgBox "Operation not completed, may you get the wrong result."
  
End Sub
    
    
Sub createNewUp()
    Application.ScreenUpdating = False
    
    'take UP file path

    Dim currentUpFilePathArr, currentUpFilePath As Variant
    currentUpFilePathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", "D:\Temp\UP Draft\Draft 2024")  ' UP file path
    If UBound(currentUpFilePathArr) = 1 Then
        currentUpFilePath = currentUpFilePathArr(1)
    Else
        MsgBox "Please select only one UP file"
        Exit Sub
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim upFolderPath, curentUpNoFromFileName As String
    upFolderPath = fso.GetParentFolderName(currentUpFilePath)
    curentUpNoFromFileName = fso.GetBaseName(currentUpFilePath)
    
    'extract UP and year of UP from file name
    Dim extractedUpAndUpYearFromFile As Variant
    extractedUpAndUpYearFromFile = Application.Run("general_utility_functions.upNoAndYearExtrac", curentUpNoFromFileName)
    
    Dim newUpFromFile As String
    newUpFromFile = extractedUpAndUpYearFromFile(1) + 1 & "/" & extractedUpAndUpYearFromFile(2)
    
    'copy current UP as new UP file
    Dim newUpFullPath As String
    newUpFullPath = upFolderPath & "\" & "UP-" & extractedUpAndUpYearFromFile(1) + 1 & "-" & extractedUpAndUpYearFromFile(2) & ".xlsx"
    
    Application.Run "general_utility_functions.CopyFileAsNewFileFSO", currentUpFilePath, newUpFullPath, True

    'take current UP no.
    Dim curentUpNo As Variant
    
    Dim newUpWb As Workbook
    Dim newUpWs As Worksheet
    Set newUpWb = Workbooks.Open(newUpFullPath)
    Set newUpWs = newUpWb.Worksheets(2)
    curentUpNo = Application.Run("helperFunctionGetData.upNoFromProvidedWs", newUpWs)
    
    'extract UP and year of UP
    Dim extractedUpAndUpYear As Variant
    extractedUpAndUpYear = Application.Run("general_utility_functions.upNoAndYearExtrac", curentUpNo)
    
    Dim newUp As String
    newUp = extractedUpAndUpYear(1) + 1 & "/" & extractedUpAndUpYear(2)
    
    If newUpFromFile <> newUp Then
        MsgBox "UP No. & UP File No. Mismatch"
        Exit Sub
    End If
    
    'change UP sheet name
    Dim newUpSheetName As String
    newUpSheetName = "UP # " & extractedUpAndUpYear(1) + 1 & "-" & extractedUpAndUpYear(2)
    newUpWs.Name = newUpSheetName
    
    'take source data as dictionary from UP Issuing Status
    Dim sourceDataAsDicUpIssuingStatus As Variant
    Set sourceDataAsDicUpIssuingStatus = Application.Run("helperFunctionGetData.sourceDataAsDicUpIssuingStatus", newUp, "UP Issuing Status for the Period # 01-03-2024 to 28-02-2025.xlsx", "UP Issuing Status # 2024-2025")
    
    Dim upNoWithWordForPutToWs, upNoInWord, yearInWord As String
    upNoInWord = Application.Run("NumToBanglaWord.numberToBanglaWord", extractedUpAndUpYear(1) + 1)
    yearInWord = Application.Run("NumToBanglaWord.numberToBanglaWord", extractedUpAndUpYear(2))
    upNoWithWordForPutToWs = newUp & " (" & upNoInWord & "/" & yearInWord & ")"
    
    newUpWs.Range("N13").value = upNoWithWordForPutToWs
    
    
    Dim upClause6RangObj As Range
    
    'previous range
    Set upClause6RangObj = Application.Run("helperFunctionGetRangeObject.upClause6BuyerinformationRangeObjectFromProvidedWs", newUpWs)
    
    'updated range
    Set upClause6RangObj = Application.Run("createUp.dealWithUpClause6", upClause6RangObj, sourceDataAsDicUpIssuingStatus)
    
    Dim upClause7RangObj As Range

    'previous range
    Set upClause7RangObj = Application.Run("helperFunctionGetRangeObject.upClause7LcinformationRangeObjectFromProvidedWs", newUpWs)
    
    'updated range
    Set upClause7RangObj = Application.Run("createUp.dealWithUpClause7", upClause7RangObj, sourceDataAsDicUpIssuingStatus)
    
    
    'working
    Application.DisplayAlerts = False
    newUpWb.Close SaveChanges:=True
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True

End Sub

Sub afterYarnConsumption()
    
'    Set nnws = Worksheets(2)

    Dim upWorkBook As Workbook
    Dim upWorksheet As Worksheet
    Dim consumptionWorksheet As Worksheet
    
    Set upWorkBook = ActiveWorkbook
    Set upWorksheet = upWorkBook.ActiveSheet 'be change sheet no. 2 & UP file
    Set consumptionWorksheet = upWorkBook.Worksheets("Consumption")
    

    Dim newUp As String
    newUp = Application.Run("helperFunctionGetData.upNoFromProvidedWs", upWorksheet)
    
    
    'take source data as dictionary from UP Issuing Status
    Dim sourceDataAsDicUpIssuingStatus As Variant
    Set sourceDataAsDicUpIssuingStatus = Application.Run("helperFunctionGetData.sourceDataAsDicUpIssuingStatus", newUp, "UP Issuing Status for the Period # 01-03-2024 to 28-02-2025.xlsx", "UP Issuing Status # 2024-2025")

    
    'take yarn consumption info from "Consumption" sheet
    Dim yarnConsumptionInfoDic As Variant
    Set yarnConsumptionInfoDic = Application.Run("afterConsumption.upYarnConsumptionInformationFromProvidedWs", consumptionWorksheet)
    
    'chemical consumption as "dedo"
    Dim finalRawMaterialsQtyDicAsGroup As Object ' SL 9 for rolling film, Fabric Qty. should be dynamic
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dedo_consumption.finalRawMaterialsQtyCalculatedAsGroup", _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "Black")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "Mercerization(Black)")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "Indigo")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "Mercerization(Indigo)")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "Topping/ Bottoming")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "Mercerization(Topping/ Bottoming)")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "Over Dying")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "Mercerization(Over Dying)")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "Coating")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "PFD")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "ECRU")), _
    yarnConsumptionInfoDic(Application.Run("general_utility_functions.RemoveInvalidChars", "TOTAL")), _
    Application.Run("utilityFunction.sumQtyFromDictFormat", sourceDataAsDicUpIssuingStatus))
    
    
    Dim impPerformanceDataDic As Object
    
    Set impPerformanceDataDic = Application.Run("data_from_imp_performance.classifiedDbDicFromImpPerformance", _
    ActiveWorkbook.path & Application.PathSeparator & "Import Performance Statement of PDL-2024-2025.xlsx") ' path change after changed the period
    
    'take source data as dictionary from Import Yarn Use Details For UD File
    Dim importYarnUseDetailsForUd As Object
    Set importYarnUseDetailsForUd = Application.Run("afterConsumption.sourceDataAsDicImportYarnUseDetailsForUd", "Import Yarn Use Details For UD.xlsx", "Use of Bill of Entry") ' path change when need

    'take UP clause 8 info from "UP" sheet
    Dim upClause8InfoDic As Object
    Set upClause8InfoDic = Application.Run("afterConsumption.upClause8InformationFromProvidedWs", upWorksheet, impPerformanceDataDic)
    
    'create new UP clause 8 info
    Dim newUpClause8InfoDic As Object
    Set newUpClause8InfoDic = Application.Run("afterConsumption.createNewUpClause8Information", upClause8InfoDic, impPerformanceDataDic, sourceDataAsDicUpIssuingStatus, importYarnUseDetailsForUd, CreateObject("Scripting.Dictionary"))
    

    
    Application.Run "afterConsumption.upClause8MakeUniqueRowsFromProvidedWs", upWorksheet
    
    Application.Run "afterConsumption.upClause8InformationPutToProvidedWs", upWorksheet, newUpClause8InfoDic
    
'    Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", Range("ab119"), finalRawMaterialsQtyDicAsGroup, 1, 1, 1 'print to sheet for test
    

End Sub
     
   
    Sub test()
        Dim test1, test2 As Variant
        
'        test1 = Worksheets("Chemicals (Import)").Range("a1").CurrentRegion.value
'        test2 = ActiveSheet.Range("a2:d2").value
'            Dim ws As Worksheet
'            Set ws = Worksheets(3)

'        Dim myDict As Object
'        Set myDict = CreateObject("Scripting.Dictionary")

        
'        Dim testReturn As Variant
'        testReturn = Application.Run("utilityFunction.SwapColumns", test1, 2, 1)
'        Set testReturn = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", myDict, "primary", Array(1, 1, 2, 1, 3))
'        Set testReturn = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", myDict, "primary", Array(8, 1, 2, 1, "dasl&"))
'        Set testReturn = Application.Run("dictionary_utility_functions.AddKeysAndValueSame", myDict, Array(8, 1, 2, 1, "dasl&"))
'        (dict As Object, mushakOrBillOfEntrySourceArr As Variant, mushakOrBillOfEntryCol As Integer, qtyCol As Integer, valueCol As Integer, discriptionCol As Integer, propertiesArr As Variant, propertiesColsArr As Variant)
'        Set testReturn = Application.Run("dictionary_utility_functions.CreateMushakOrBillOfEntryDbDict", myDict, test1, 3, 8, 9, 7, Array("qty", "value", "discription", "usedQty", "usedValue"), Array(8, 9, 7, 10, 11))
        
'        Set testReturn = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", Array("LC", "q t+ y", "va  lue"), Array(8230, 10, 20))
'        ActiveSheet.Range("a10:i14").value = testReturn

'        Dim ws As Worksheet
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
'
'        Dim i As LongLong
'        For i = 1 To 10
'            dict("key_" & dict.Count + 1) = "Value_" & i
'        Next i
'

'        Set test1 = Application.Run("utilityFunction.CombinedAllSheetsMushakOrBillOfEntryDbDict", "D:\Temp\UP Draft\Draft 2024\Import Performance Statement of PDL-2024-2025.xlsx")
'
'        Set ws = ActiveSheet
'        ws.Cells.Clear
'        test2 = testReturn.keys
'        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a1"), test1("C89828121052022_17214_72127"), 1, 1, 0
'        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a3"), test1("C134115724082023_10966_33994"), 1, 1, 0
'        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a5"), test1("M632703304092023_11650_20387"), 1, 1, 0
'        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a7"), test1("C108682713072023_20000_25000"), 1, 1, 0
'        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a9"), test1("C127315612082023_23040_32256"), 1, 1, 0
'        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a11"), test1("M633428402092023_10020_4509"), 1, 1, 0
'        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a13"), test1("C59994513042023_24000_69600"), 1, 1, 0
'
'        test1 = Application.Run("general_utility_functions.ExtractLeftDigitWithRegex", "56546544654.4454")
'
'        test1 = Application.Run("general_utility_functions.dictKeyGeneratorWithMushakOrBillOfEntryQtyAndValue", "M daslk3 f", 5464.2547, "56546544654.4454")
'

    
'    Application.Run "utility_formating_fun.borderInsideHairlineAroundThin", Selection

'    Set dict = Application.Run("general_utility_functions.sequentiallyRelateTwoArraysAsDictionary", "ip", "date", Array("IP1 &_"), Array("10/12/2024"))
    
'    test1 = Application.Run("general_utility_functions.isStrPatternExist", "abc", ".", True, True, True)

'    Set test1 = Application.Run("general_utility_functions.regExReturnedObj", "a58b6c", "\d+", True, True, True)
    
    test1 = Application.Run("general_utility_functions.extractAndFormatUdNo", "BGMEA/DHK/UD/2024/3578/020")
    test1 = Application.Run("general_utility_functions.extractAndFormatUdNo", "BGMEA/DHK/AM/2024/3016/002-003")

'    Set test1 = Application.Run("dedo_consumption.ropeDyingBlackPretreatmentAndDying")

'    Set test1 = Application.Run("dedo_consumption.ropeDyingBlackPretreatmentAndDying")
'    Set test2 = Application.Run("dedo_consumption.wtpWaterTreatmentPlant")
'
'    Set dict = Application.Run("dictionary_utility_functions.mergeDict", test1, test2)

'    Set dict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", dict, "daslkj ", 654)
    
'    test1 = Array( _
'    "Sulphur Black (Powder)  or_Sl_5", _
'    "Vat Dyes (Liquid)_Sl_26", _
'    "Vat Dyes (Powder/Solid)  or_Sl_44", _
'    "Sulphur Black (Powder)  or_Sl_46", _
'    "Paper Tube_Sl_85" _
'    )
    Set dict = Application.Run("dedo_consumption.combineAllDedoConDicAfterCalculateActualQty", 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100)
'
'    test2 = Application.Run("dictionary_utility_functions.sumOfProvidedKeys", dict, test1)
'
'
'    Set dict = Application.Run("dedo_consumption.appliedNotUsedRawMaterials", dict, test1)
'
'    test2 = Application.Run("dictionary_utility_functions.sumOfProvidedKeys", dict, test1)
'
    Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", Range("a1"), dict, 1, 1, 1

    Set dict = Application.Run("dedo_consumption.finalRawMaterialsQtyCalculatedAsGroup", 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100)

    
'    Application.Run "afterConsumption.upClause8MakeUniqueRowsFromProvidedWs", ActiveSheet
    
    End Sub
    

