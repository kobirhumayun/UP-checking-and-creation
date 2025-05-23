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
sourceDataUpIssuingStatus = Application.Run("helperFunctionGetData.SourceDataUPIssuingStatus", upNo, "UP Issuing Status for the Period # 01-03-2025 to 28-02-2026.xlsx", "UP Issuing Status # 2025-2026")

'take source data as dictionary from UP Issuing Status
Dim sourceDataAsDicUpIssuingStatus As Variant
Set sourceDataAsDicUpIssuingStatus = Application.Run("helperFunctionGetData.sourceDataAsDicUpIssuingStatus", upNo, "UP Issuing Status for the Period # 01-03-2025 to 28-02-2026.xlsx", "UP Issuing Status # 2025-2026")


'take source data from Import Performance Yarn Import
Dim sourceDataImportPerformanceYarnImport As Variant
sourceDataImportPerformanceYarnImport = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2025-2026.xlsx", "Yarn (Import)", True, False)


'take source data from Import Performance Yarn Local
Dim sourceDataImportPerformanceYarnLocal As Variant
sourceDataImportPerformanceYarnLocal = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2025-2026.xlsx", "Yarn (Local)", False, False)


'take source data from Import Performance Dyes
Dim sourceDataImportPerformanceDyes As Variant
sourceDataImportPerformanceDyes = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2025-2026.xlsx", "Dyes", False, False)


'take source data from Import Performance Chemicals Import
Dim sourceDataImportPerformanceChemicalsImport As Variant
sourceDataImportPerformanceChemicalsImport = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2025-2026.xlsx", "Chemicals (Import)", False, False)


'take source data from Import Performance Chemicals Local
Dim sourceDataImportPerformanceChemicalsLocal As Variant
sourceDataImportPerformanceChemicalsLocal = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2025-2026.xlsx", "Chemicals (Local)", False, False)


'take source data from Import Performance Stretch Wrapping Film
Dim sourceDataImportPerformanceStretchWrappingFilm As Variant
sourceDataImportPerformanceStretchWrappingFilm = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2025-2026.xlsx", "St.Wrap.Film (Import)", False, False)


'take source data from Import Performance Total Summary
Dim sourceDataImportPerformanceTotalSummary As Variant
sourceDataImportPerformanceTotalSummary = Application.Run("helperFunctionGetData.sourceDataImportPerformance", "Import Performance Statement of PDL-2025-2026.xlsx", "Summary of Grand Total", False, True)



'take source data from previous UP clause 8
Dim sourceDataPreviousUpClause8 As Variant
sourceDataPreviousUpClause8 = Application.Run("helperFunctionGetData.sourceDataPreviousUp", upNo, 8, True, False)



'take source data from previous UP clause 9
Dim sourceDataPreviousUpClause9 As Variant
sourceDataPreviousUpClause9 = Application.Run("helperFunctionGetData.sourceDataPreviousUp", upNo, 9, False, True)






'#####compare section start from here#####

Dim upClause6And7CompareWithSourceResultArr As Variant
upClause6And7CompareWithSourceResultArr = Application.Run("helperFunctionCompareData.upClause6And7CompareWithSource", upClause6BuyerinformationRangeObject, upClause7LcinformationRangeObject, sourceDataUpIssuingStatus, sourceDataAsDicUpIssuingStatus)


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

Sub totalPeriodBillOfEntryOrMushkUsedCalculationAndPutToImportPerformanceWithJson()

  Application.ScreenUpdating = False
  
  On Error GoTo ErrorMsg

    Dim importPerformanceFileWb As Workbook
    Set importPerformanceFileWb = ActiveWorkbook
        
    Dim allUpClause8UseAsMushakOrBillOfEntryDic As Object
    Set allUpClause8UseAsMushakOrBillOfEntryDic = Application.Run("general_utility_functions.sumUsedQtyAndValueAsMushakOrBillOfEntryFromSelectedUpFile")

    Dim upSequenceStr As String
    upSequenceStr = Application.Run("utilityFunction.upSequenceStrGenerator", allUpClause8UseAsMushakOrBillOfEntryDic("allCalculatedUpList").keys, " -to- ", 10)

    Dim importPerformanceFileYarnImportWs As Worksheet
    Set importPerformanceFileYarnImportWs = importPerformanceFileWb.Worksheets("Yarn (Import)")

    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFileWithJson", importPerformanceFileYarnImportWs, 3, 4, 7, 8, 9, 10, 28, allUpClause8UseAsMushakOrBillOfEntryDic     'Used Qty & Value put to import performance file
             
    Dim importPerformanceFileYarnLocalWs As Worksheet
    Set importPerformanceFileYarnLocalWs = importPerformanceFileWb.Worksheets("Yarn (Local)")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFileWithJson", importPerformanceFileYarnLocalWs, 3, 4, 7, 8, 9, 10, 28, allUpClause8UseAsMushakOrBillOfEntryDic     'Used Qty & Value put to import performance file
        
    Dim importPerformanceFileDyesWs As Worksheet
    Set importPerformanceFileDyesWs = importPerformanceFileWb.Worksheets("Dyes")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFileWithJson", importPerformanceFileDyesWs, 3, 4, 7, 8, 9, 10, 28, allUpClause8UseAsMushakOrBillOfEntryDic     'Used Qty & Value put to import performance file
        
    Dim importPerformanceFileChemialsImportWs As Worksheet
    Set importPerformanceFileChemialsImportWs = importPerformanceFileWb.Worksheets("Chemicals (Import)")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFileWithJson", importPerformanceFileChemialsImportWs, 3, 4, 8, 9, 10, 11, 28, allUpClause8UseAsMushakOrBillOfEntryDic     'Used Qty & Value put to import performance file
        
    Dim importPerformanceFileChemialsLocalWs As Worksheet
    Set importPerformanceFileChemialsLocalWs = importPerformanceFileWb.Worksheets("Chemicals (Local)")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFileWithJson", importPerformanceFileChemialsLocalWs, 3, 4, 8, 9, 10, 11, 28, allUpClause8UseAsMushakOrBillOfEntryDic     'Used Qty & Value put to import performance file
        
    Dim importPerformanceFileWrappingFilmWs As Worksheet
    Set importPerformanceFileWrappingFilmWs = importPerformanceFileWb.Worksheets("St.Wrap.Film (Import)")
    
    Application.Run "utilityFunction.putTotalUsedQtyAndValueAsBillOfEntryOrMushakToImportPerformanceFileWithJson", importPerformanceFileWrappingFilmWs, 3, 4, 8, 9, 10, 11, 28, allUpClause8UseAsMushakOrBillOfEntryDic     'Used Qty & Value put to import performance file

    MsgBox upSequenceStr
    
  Application.ScreenUpdating = True
  
  Exit Sub
  
ErrorMsg:
  MsgBox "Operation not completed, may you get the wrong result."
  
End Sub
    
    
Sub createNewUp()
    Application.ScreenUpdating = False

    Dim answer As VbMsgBoxResult

            ' Display the message box with Yes and No buttons
        answer = MsgBox("Do you want make complete UP? If click to No button then stop before yarn consumption.", vbYesNo + vbQuestion, "Create UP")
    
    'take UP file path

    Dim currentUpFilePathArr, currentUpFilePath As Variant
    currentUpFilePathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", "D:\Temp\UP Draft\Draft 2025")  ' UP file path
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
    Dim extractedUpAndUpYearFromFile As Object
    Set extractedUpAndUpYearFromFile = Application.Run("general_utility_functions.upNoAndYearExtracAsDict", curentUpNoFromFileName)

    Dim importPerformanceFileName As String
    importPerformanceFileName = "Import Performance Statement of PDL-2025-2026.xlsx"

    'take source data from Import Performance dyes to check last UP updated or not
    Dim sourceDataImportPerformanceDyes As Variant
    sourceDataImportPerformanceDyes = Application.Run("helperFunctionGetData.sourceDataImportPerformanceWithUpColumn", upFolderPath & Application.PathSeparator & importPerformanceFileName, "Dyes", True, True)

    Dim isLastUpUsedUpdated As Boolean
    isLastUpUsedUpdated = Application.Run("afterConsumption.isLastUpUsedUpdatedInImportPerformance", _
        sourceDataImportPerformanceDyes, extractedUpAndUpYearFromFile("only_up_no") & "/" & extractedUpAndUpYearFromFile("only_up_year"), 28)

    If Not isLastUpUsedUpdated Then
        MsgBox "Current UP-" & extractedUpAndUpYearFromFile("only_up_no") & "/" & extractedUpAndUpYearFromFile("only_up_year") & _
        " not updated as last UP in import performance for used Bill of Entry or Mushak" & Chr(10) & "Update first!"
        Exit Sub
    End If
    
    Dim newUpFromFile As String
    Dim newUpOnlyFromFile As String
    
    If extractedUpAndUpYearFromFile("only_up_no") < 9 Then
        newUpOnlyFromFile = "0" & extractedUpAndUpYearFromFile("only_up_no") + 1
    Else
        newUpOnlyFromFile = extractedUpAndUpYearFromFile("only_up_no") + 1
    End If
    
    newUpFromFile = newUpOnlyFromFile & "/" & extractedUpAndUpYearFromFile("only_up_year")

    'copy current UP as new UP file
    Dim newUpFullPath As String
    newUpFullPath = upFolderPath & "\" & "UP-" & newUpOnlyFromFile & "-" & extractedUpAndUpYearFromFile("only_up_year") & ".xlsx"
    
    Application.Run "general_utility_functions.CopyFileAsNewFileFSO", currentUpFilePath, newUpFullPath, True

    'take current UP no.
    Dim curentUpNo As Variant
    
    Dim newUpWb As Workbook
    Dim newUpWs As Worksheet
    Set newUpWb = Workbooks.Open(newUpFullPath)
    Set newUpWs = newUpWb.Worksheets(2)
    curentUpNo = Application.Run("helperFunctionGetData.upNoFromProvidedWs", newUpWs)
    
    'extract UP and year of UP
    Dim extractedUpAndUpYear As Object
    Set extractedUpAndUpYear = Application.Run("general_utility_functions.upNoAndYearExtracAsDict", curentUpNo)
    
    Dim newUp As String
    Dim newUpOnly As String
    
    If extractedUpAndUpYear("only_up_no") < 9 Then
        newUpOnly = "0" & extractedUpAndUpYear("only_up_no") + 1
    Else
        newUpOnly = extractedUpAndUpYear("only_up_no") + 1
    End If

    newUp = newUpOnly & "/" & extractedUpAndUpYear("only_up_year")
    
    If newUpFromFile <> newUp Then
        MsgBox "UP No. & UP File No. Mismatch"
        Exit Sub
    End If
    
    'change UP sheet name
    Dim newUpSheetName As String
    newUpSheetName = "UP # " & newUpOnly & "-" & extractedUpAndUpYear("only_up_year")
    newUpWs.Name = newUpSheetName
    
    'take source data as dictionary from UP Issuing Status
    Dim sourceDataAsDicUpIssuingStatus As Variant
    Set sourceDataAsDicUpIssuingStatus = Application.Run("helperFunctionGetData.sourceDataAsDicUpIssuingStatus", newUp, "UP Issuing Status for the Period # 01-03-2025 to 28-02-2026.xlsx", "UP Issuing Status # 2025-2026")
    
    Dim upNoWithWordForPutToWs, upNoInWord, yearInWord As String
    upNoInWord = Application.Run("NumToBanglaWord.numberToBanglaWord", newUpOnly)
    yearInWord = Application.Run("NumToBanglaWord.numberToBanglaWord", extractedUpAndUpYear("only_up_year"))
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

        ' Check which button the user clicked
    If answer = vbYes Then
            ' Code to execute if user clicks Yes

        newUpWb.Activate

        Application.Run "main.dealWithYarnConsumption"
        Application.Run "main.afterYarnConsumption"
        Application.Run "main.dealWithNote"

        Application.DisplayAlerts = False
        newUpWb.Close SaveChanges:=True
        Application.DisplayAlerts = True

    ElseIf answer = vbNo Then
            ' Code to execute if user clicks No

        Application.DisplayAlerts = False
        newUpWb.Close SaveChanges:=True
        Application.DisplayAlerts = True

    End If
    
    Application.ScreenUpdating = True

End Sub

Sub afterYarnConsumption()

    Application.ScreenUpdating = False

    Dim upWorkBook As Workbook
    Dim upWorksheet As Worksheet
    Dim consumptionWorksheet As Worksheet
    
    Set upWorkBook = ActiveWorkbook
    Set upWorksheet = upWorkBook.Worksheets(2)
    Set consumptionWorksheet = upWorkBook.Worksheets("Consumption")
    

    Dim newUp As String
    newUp = Application.Run("helperFunctionGetData.upNoFromProvidedWs", upWorksheet)
    
    
    'take source data as dictionary from UP Issuing Status
    Dim sourceDataAsDicUpIssuingStatus As Variant
    Set sourceDataAsDicUpIssuingStatus = Application.Run("helperFunctionGetData.sourceDataAsDicUpIssuingStatus", newUp, "UP Issuing Status for the Period # 01-03-2025 to 28-02-2026.xlsx", "UP Issuing Status # 2025-2026")

    
    'take yarn consumption info from "Consumption" sheet
    Dim yarnConsumptionInfoDic As Variant
    Set yarnConsumptionInfoDic = Application.Run("afterConsumption.upYarnConsumptionInformationFromProvidedWs", consumptionWorksheet)
    
    'chemical consumption as "dedo"
    Dim finalRawMaterialsQtyDicAsGroup As Object
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
    
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "cotton", yarnConsumptionInfoDic("Cotton"))
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "polyester", yarnConsumptionInfoDic("Polyester"))
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "spandex", yarnConsumptionInfoDic("Spandex"))
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "Detergent", 0) ' Qty. be dynamic
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "Pumice Stone", 0) ' Qty. be dynamic
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "Natural Garnet", 0) ' Qty. be dynamic
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "Hydroxylamine", 0) ' Qty. be dynamic
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "Bleaching Powder", 0) ' Qty. be dynamic
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "Finishing Agent", 0) ' Qty. be dynamic
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "Antistain", 0) ' Qty. be dynamic
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "Reactive Dyes", 0) ' Qty. be dynamic
    Set finalRawMaterialsQtyDicAsGroup = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", finalRawMaterialsQtyDicAsGroup, "Polymer", 0) ' Qty. be dynamic
    
    Dim impPerformanceDataDic As Object
    Dim importPerformanceFileName As String
    importPerformanceFileName = "Import Performance Statement of PDL-2025-2026.xlsx"

    'take source data from Import Performance Total Summary
    Dim sourceDataImportPerformanceTotalSummary As Variant
    sourceDataImportPerformanceTotalSummary = Application.Run("helperFunctionGetData.sourceDataImportPerformance", importPerformanceFileName, "Summary of Grand Total", True, False)

    Set impPerformanceDataDic = Application.Run("data_from_imp_performance.classifiedDbDicFromImpPerformance", _
    ActiveWorkbook.path & Application.PathSeparator & importPerformanceFileName) ' path change after changed the period
    
    'take source data as dictionary from Import Yarn Use Details For UD File
    Dim importYarnUseDetailsForUd As Object
    Set importYarnUseDetailsForUd = Application.Run("afterConsumption.sourceDataAsDicImportYarnUseDetailsForUd", "Import Yarn Use Details For UD of 2025-2026.xlsx", "Use of Bill of Entry") ' path change when need

    'take UP clause 8 info from "UP" sheet
    Dim upClause8InfoDic As Object
    Set upClause8InfoDic = Application.Run("afterConsumption.upClause8InformationForCreateUpFromProvidedWs", upWorksheet, impPerformanceDataDic)
    
    'create new UP clause 8 info
    Dim newUpClause8InfoDic As Object
    Set newUpClause8InfoDic = Application.Run("afterConsumption.createNewUpClause8Information", upClause8InfoDic, impPerformanceDataDic, sourceDataAsDicUpIssuingStatus, importYarnUseDetailsForUd, finalRawMaterialsQtyDicAsGroup)

    'create new UP clause 8 yarn, dyes chemicals Classified part Qty. & value
    Dim newUpClause8InfoClassifiedPartDic As Object
    Set newUpClause8InfoClassifiedPartDic = Application.Run("afterConsumption.sumNewUpClause8ClassifiedPart", newUpClause8InfoDic)

    'add consumption range to UP issuing status
    Dim withConRangeSourceDataAsDicUpIssuingStatus As Object
    Set withConRangeSourceDataAsDicUpIssuingStatus = Application.Run("afterConsumption.addConRangeToSourceDataAsDicUpIssuingStatus", consumptionWorksheet, sourceDataAsDicUpIssuingStatus)
    

    
    Application.Run "afterConsumption.upClause8MakeUniqueRowsFromProvidedWs", upWorksheet
    
    Application.Run "afterConsumption.upClause8InformationPutToProvidedWs", upWorksheet, newUpClause8InfoDic

    Application.Run "afterConsumption.dealWithUpClause9", upWorksheet, newUpClause8InfoClassifiedPartDic, sourceDataImportPerformanceTotalSummary

    Application.Run "afterConsumption.dealWithUpClause11", upWorksheet, withConRangeSourceDataAsDicUpIssuingStatus

    Application.Run "afterConsumption.dealWithUpClause12a", upWorksheet, withConRangeSourceDataAsDicUpIssuingStatus

    Application.Run "afterConsumption.dealWithUpClause12b", upWorksheet, sourceDataAsDicUpIssuingStatus
    
    Application.Run "afterConsumption.dealWithUpClause12bGarments", upWorksheet

    Application.Run "afterConsumption.dealWithUpClause13", upWorksheet, newUpClause8InfoClassifiedPartDic

    Application.Run "afterConsumption.dealWithUpClause14", upWorksheet, sourceDataAsDicUpIssuingStatus

    With upWorksheet.Cells
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlAutomatic
    End With

    Application.ScreenUpdating = True

End Sub

Sub updateAfterUpClause8()

    Application.ScreenUpdating = False

    Dim upWorkBook As Workbook
    Dim upWorksheet As Worksheet
    Dim consumptionWorksheet As Worksheet
    
    Set upWorkBook = ActiveWorkbook
    Set upWorksheet = upWorkBook.Worksheets(2)
    Set consumptionWorksheet = upWorkBook.Worksheets("Consumption")

    Dim newUp As String
    newUp = Application.Run("helperFunctionGetData.upNoFromProvidedWs", upWorksheet)
    
    'take source data as dictionary from UP Issuing Status
    Dim sourceDataAsDicUpIssuingStatus As Variant
    Set sourceDataAsDicUpIssuingStatus = Application.Run("helperFunctionGetData.sourceDataAsDicUpIssuingStatus", newUp, "UP Issuing Status for the Period # 01-03-2025 to 28-02-2026.xlsx", "UP Issuing Status # 2025-2026")

    Dim extractedUpAndUpYear As Object
    Set extractedUpAndUpYear = Application.Run("general_utility_functions.upNoAndYearExtracAsDict", newUp)

    Dim previousUpOnlyNo As String
    previousUpOnlyNo = extractedUpAndUpYear("only_up_no") - 1
    
    If previousUpOnlyNo < 10 Then
        previousUpOnlyNo = "0" & previousUpOnlyNo
    End If
    
    Dim previousUpfileName As String
    previousUpfileName = "UP-" & previousUpOnlyNo & "-" & extractedUpAndUpYear("only_up_year") & ".xlsx"

    Dim previousUpClause9Info As Variant
    previousUpClause9Info = Application.Run("afterConsumption.upClause9InfoFromProvidedFile", previousUpfileName, True, True)

    Dim importPerformanceFileName As String
    importPerformanceFileName = "Import Performance Statement of PDL-2025-2026.xlsx"

    'take source data from Import Performance Total Summary
    Dim sourceDataImportPerformanceTotalSummary As Variant
    sourceDataImportPerformanceTotalSummary = Application.Run("helperFunctionGetData.sourceDataImportPerformance", importPerformanceFileName, "Summary of Grand Total", True, True)

    'take UP clause 8 info from "UP" sheet
    Dim upClause8InfoDic As Object
    Set upClause8InfoDic = Application.Run("general_utility_functions.upClause8InformationFromProvidedWs", upWorksheet)

    'create UP clause 8 yarn, dyes chemicals Classified part Qty. & value
    Dim upClause8InfoClassifiedPartDic As Object
    Set upClause8InfoClassifiedPartDic = Application.Run("afterConsumption.sumUpClause8ClassifiedPart", upClause8InfoDic)

    'add consumption range to UP issuing status
    Dim withConRangeSourceDataAsDicUpIssuingStatus As Object
    Set withConRangeSourceDataAsDicUpIssuingStatus = Application.Run("afterConsumption.addConRangeToSourceDataAsDicUpIssuingStatus", consumptionWorksheet, sourceDataAsDicUpIssuingStatus)
    
    Application.Run "afterConsumption.dealWithUpClause9WithPreviousUpData", upWorksheet, upClause8InfoClassifiedPartDic, sourceDataImportPerformanceTotalSummary, previousUpClause9Info

    Application.Run "afterConsumption.dealWithUpClause11", upWorksheet, withConRangeSourceDataAsDicUpIssuingStatus

    Application.Run "afterConsumption.dealWithUpClause12a", upWorksheet, withConRangeSourceDataAsDicUpIssuingStatus

    Application.Run "afterConsumption.dealWithUpClause12b", upWorksheet, sourceDataAsDicUpIssuingStatus

    Application.Run "afterConsumption.dealWithUpClause13", upWorksheet, upClause8InfoClassifiedPartDic

    Application.Run "afterConsumption.dealWithUpClause14", upWorksheet, sourceDataAsDicUpIssuingStatus

    With upWorksheet.Cells
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlAutomatic
    End With

    Application.ScreenUpdating = True

    MsgBox "UP " & newUp & " clause 9-13 updated!"

End Sub

Sub dealWithYarnConsumption()

    Dim upWorkBook As Workbook
    Dim upWorksheet As Worksheet
    Dim consumptionWorksheet As Worksheet

    Set upWorkBook = ActiveWorkbook
    Set upWorksheet = upWorkBook.Worksheets(2)
    Set consumptionWorksheet = upWorkBook.Worksheets("Consumption")


    Dim newUp As String
    newUp = Application.Run("helperFunctionGetData.upNoFromProvidedWs", upWorksheet)

    Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
    Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

    consumptionWorksheet.Range("a1").value = vsCodeNotSupportedOrBengaliTxtDictionary("pioneerDenimLimitedUpNoBengaliTxt") & newUp

        'take source data as dictionary from UP Issuing Status
    Dim sourceDataAsDicUpIssuingStatus As Variant
    Set sourceDataAsDicUpIssuingStatus = Application.Run("helperFunctionGetData.sourceDataAsDicUpIssuingStatus", newUp, "UP Issuing Status for the Period # 01-03-2025 to 28-02-2026.xlsx", "UP Issuing Status # 2025-2026")

        'add PI info to UP Issuing Status
    Dim withPiInfosourceDataAsDicUpIssuingStatus As Variant
    Set withPiInfosourceDataAsDicUpIssuingStatus = Application.Run("yarnConsumption.addPiInfoSourceDataAsDicUpIssuingStatus", sourceDataAsDicUpIssuingStatus)

        'validate commercial file qty. & unit with PI info
    Application.Run "yarnConsumption.validateCommercialFileQtyAndUnit", withPiInfosourceDataAsDicUpIssuingStatus

        'add yarn consumption info to UP Issuing Status
    Dim withYarnConsumptionInfosourceDataAsDicUpIssuingStatus As Variant
    Set withYarnConsumptionInfosourceDataAsDicUpIssuingStatus = Application.Run("yarnConsumption.addYarnConsumptionInfoSourceDataAsDicUpIssuingStatus", withPiInfosourceDataAsDicUpIssuingStatus)

    Application.Run "yarnConsumption.dealWithConsumptionSheet", consumptionWorksheet, withYarnConsumptionInfosourceDataAsDicUpIssuingStatus


     
End Sub

Sub dealWithNote()

    Application.ScreenUpdating = False

    Dim upWorkBook As Workbook
    Dim upWorksheet As Worksheet
    Dim consumptionWorksheet As Worksheet
    Dim noteWorksheet As Worksheet
    
    Set upWorkBook = ActiveWorkbook
    Set upWorksheet = upWorkBook.Worksheets(2)
    Set consumptionWorksheet = upWorkBook.Worksheets("Consumption")
    Set noteWorksheet = upWorkBook.Worksheets("Note")

    Dim newUp As String
    newUp = Application.Run("helperFunctionGetData.upNoFromProvidedWs", upWorksheet)
    
    'take source data as dictionary from UP Issuing Status
    Dim sourceDataAsDicUpIssuingStatus As Variant
    Set sourceDataAsDicUpIssuingStatus = Application.Run("helperFunctionGetData.sourceDataAsDicUpIssuingStatus", newUp, "UP Issuing Status for the Period # 01-03-2025 to 28-02-2026.xlsx", "UP Issuing Status # 2025-2026")

    'take UP clause 8 info from "UP" sheet
    Dim upClause8InfoDic As Object
    Set upClause8InfoDic = Application.Run("general_utility_functions.upClause8InformationFromProvidedWs", upWorksheet)

    'create UP clause 8 yarn, dyes chemicals Classified part Qty. & value
    Dim upClause8InfoClassifiedPartDic As Object
    Set upClause8InfoClassifiedPartDic = Application.Run("afterConsumption.sumUpClause8ClassifiedPart", upClause8InfoDic)

    'add consumption range to UP issuing status
    Dim withConRangeSourceDataAsDicUpIssuingStatus As Object
    Set withConRangeSourceDataAsDicUpIssuingStatus = Application.Run("afterConsumption.addConRangeToSourceDataAsDicUpIssuingStatus", consumptionWorksheet, sourceDataAsDicUpIssuingStatus)
    
    Application.Run "upNote.putUpSummary", noteWorksheet, sourceDataAsDicUpIssuingStatus, upClause8InfoClassifiedPartDic, newUp
    Application.Run "upNote.putLcInfo", noteWorksheet, sourceDataAsDicUpIssuingStatus
    Application.Run "upNote.putUdIpExpInfo", noteWorksheet, sourceDataAsDicUpIssuingStatus
    Application.Run "upNote.putBuyerAndBankInfo", noteWorksheet, sourceDataAsDicUpIssuingStatus
    ' Application.Run "upNote.putVerifiedInfo", noteWorksheet, sourceDataAsDicUpIssuingStatus 'not submitted now
    Application.Run "upNote.putRawMaterialsQtyAsGroup", noteWorksheet, upClause8InfoDic

    Dim upClause8BtbLcinformationRangeObject As Object
    Set upClause8BtbLcinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause8BtbLcinformationRangeObjectFromProvidedWs", upWorksheet)

    With upWorksheet.Cells
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlAutomatic
    End With

    Application.Run "utilityFunction.cellsMarkingAsValue", upClause8BtbLcinformationRangeObject.Range("G1:G" & upClause8BtbLcinformationRangeObject.Rows.Count), "Mushak Pending"

    Application.ScreenUpdating = True

    MsgBox "UP " & newUp & " making done!"

End Sub

Sub createExportImportPerformanceAsUp()

    Dim totalUpListForReport As Variant
    totalUpListForReport = Application.Run("general_utility_functions.upSequenceArrayFromUpRange")
    
    Dim jsonPathArr As Variant

    jsonPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", "D:\Temp\UP Draft\Draft 2025\json-all-up-clause")  ' JSON file path

    If Not UBound(jsonPathArr) = 1 Then
        MsgBox "Please select only one JSON file"
        Exit Sub
    End If
    
    Dim allUpDicFromJson As Object
    Set allUpDicFromJson = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", jsonPathArr(1))
    
    Dim basePath As String
    basePath = "D:\Temp\UP Draft\Draft 2025\Import & Export Performace 2024-2025"
    
    Dim sampleUpFilePathDeem As String
    Dim sampleUpFilePathDirect As String
    
    sampleUpFilePathDeem = basePath & Application.PathSeparator & "Import-Export-UP-Performance-Deem-Sample.xlsx"
    sampleUpFilePathDirect = basePath & Application.PathSeparator & "Import-Export-UP-Performance-Direct-Sample.xlsx"

    Dim upNoAndDtFilePath As String
    upNoAndDtFilePath = basePath & Application.PathSeparator & "_up-no-and-date.xlsx"

    Dim upNoAndDtAsDict As Object
    Set upNoAndDtAsDict = Application.Run("reportAsUp.upNoAndDtAsDict", upNoAndDtFilePath)
    
        'act like middlewareFunction
    Dim isExistRelatedUpDate As Boolean
    isExistRelatedUpDate = Application.Run("reportAsUp.isExistRelatedUpDate", upNoAndDtAsDict, totalUpListForReport)
    
    If Not isExistRelatedUpDate Then
        Exit Sub
    End If

    Dim newReportFilesPath As Object
    Set newReportFilesPath = Application.Run("reportAsUp.copySmpleFileAsNewReportFileAndReturnAllPath", basePath, _
        sampleUpFilePathDeem, sampleUpFilePathDirect, totalUpListForReport, allUpDicFromJson)
    
    Application.Run "reportAsUp.putValueToReportDeemUp", allUpDicFromJson, newReportFilesPath("deemUpFullPathDict"), upNoAndDtAsDict
    Application.Run "reportAsUp.putValueToReportDirectUp", allUpDicFromJson, newReportFilesPath("directUpFullPathDict"), upNoAndDtAsDict

    MsgBox "Required UP created."
    
End Sub

Sub CreateRawMaterialsGroupReportAsUp()

    Dim jsonPathArr As Variant

    jsonPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", "D:\Temp\UP Draft\Draft 2025\json-all-up-clause")  ' JSON file path

    If Not UBound(jsonPathArr) = 1 Then
        MsgBox "Please select only one JSON file"
        Exit Sub
    End If
    
    Dim allUpDicFromJson As Object
    Set allUpDicFromJson = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", jsonPathArr(1))
    
    Dim basePath As String
    basePath = "D:\Temp\UP Draft\Draft 2025\up-raw-material-report"

    Dim groupBydictionaries As Object
    Set groupBydictionaries = CreateObject("Scripting.Dictionary")

    Dim totalUniqueGroupName As Object
    Set totalUniqueGroupName = CreateObject("Scripting.Dictionary")

    Dim temp As Object

     Dim currentKey As Variant
    
    ' Iterate through the input dictionary
    For Each currentKey In allUpDicFromJson.keys

      Set temp = Application.Run("reportAsUp.GroupByKeyAndSum", allUpDicFromJson(currentKey)("upClause8"), "nameOfGoods", "inThisUpUsedQtyOfGoods")

        If Not groupBydictionaries.Exists(currentKey) Then
            groupBydictionaries.Add currentKey, temp
        End If

        Dim currentKeyTemp As Variant
        For Each currentKeyTemp In temp.keys
            If Not totalUniqueGroupName.Exists(currentKeyTemp) Then
                totalUniqueGroupName.Add currentKeyTemp, currentKeyTemp
            End If
        Next currentKeyTemp
        
    Next currentKey

    Dim sortedAllCalculatedUp As Variant
    sortedAllCalculatedUp = Application.Run("Sorting_Algorithms.upSort", groupBydictionaries.keys)

    groupBydictionaries.Add "totalUniqueGroupName", totalUniqueGroupName

    Application.Run "JsonUtilityFunction.SaveDictionaryToJsonTextFile", groupBydictionaries, basePath & Application.PathSeparator & _
        "UP-" & Replace(sortedAllCalculatedUp(LBound(sortedAllCalculatedUp)), "/", "-") & "-to-" & _
        Replace(sortedAllCalculatedUp(UBound(sortedAllCalculatedUp)), "/", "-") & "-clause8-group-by-raw-materials-data" & ".json"

    Dim formatedGroupedDictionaryAsReportWs As Object
    Set formatedGroupedDictionaryAsReportWs = Application.Run("reportAsUp.GroupedDictionaryFormateAsReportWs", groupBydictionaries)

    Application.Run "JsonUtilityFunction.SaveDictionaryToJsonTextFile", formatedGroupedDictionaryAsReportWs, basePath & Application.PathSeparator & _
       "-formatedGroupedDictionaryAsReportWs" & ".json"
    
    Application.Run "reportAsUp.PutRawMaterialsGroupDataToWs", formatedGroupedDictionaryAsReportWs, ActiveSheet

End Sub

Sub CreateQuantityOfGoodsUsedInProductionReportAsUp()

    Dim jsonPathArr As Variant

    jsonPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", "D:\Temp\UP Draft\Draft 2025\json-all-up-clause")  ' JSON file path

    If Not UBound(jsonPathArr) = 1 Then
        MsgBox "Please select only one JSON file"
        Exit Sub
    End If
    
    Dim allUpDicFromJson As Object
    Set allUpDicFromJson = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", jsonPathArr(1))
    
    Application.Run "reportAsUp.PutQuantityOfGoodsUsedInProductionDataToWs", allUpDicFromJson, ActiveSheet

End Sub

Sub ObjectJsonFileSaveAsArrayOfObjectJson()

    Dim jsonPathArr As Variant
    jsonPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", "D:\Temp\UP Draft\Draft 2025\json-all-up-clause")  ' JSON file path

    If Not UBound(jsonPathArr) = 1 Then
        MsgBox "Please select only one JSON file"
        Exit Sub
    End If

    Dim dicFromJson As Object
    Set dicFromJson = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", jsonPathArr(1))

    Dim convertedArrOfDict As Variant
    convertedArrOfDict = Application.Run("dictionary_utility_functions.ConvertDictToArrayOfDict", dicFromJson)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderPath, curentFileName As String
    folderPath = fso.GetParentFolderName(jsonPathArr(1))
    curentFileName = fso.GetFileName(jsonPathArr(1))

    if Not fso.FolderExists(folderPath & Application.PathSeparator & "array-of-object") Then
        fso.CreateFolder folderPath & Application.PathSeparator & "array-of-object"
    End If

    folderPath = folderPath & Application.PathSeparator & "array-of-object"

    Dim newJsonPath As String
    newJsonPath = folderPath & Application.PathSeparator & "array-of-object-" & curentFileName

    Application.Run "JsonUtilityFunction.SaveArrayOfDictionaryToJsonTextFile", convertedArrOfDict, newJsonPath

    MsgBox "Array of object JSON file created."

End Sub

    Sub test()
        Dim test1, test2 As Variant
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        
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
'
'        Dim i As LongLong
'        For i = 1 To 10
'            dict("key_" & dict.Count + 1) = "Value_" & i
'        Next i
'

'        Set test1 = Application.Run("utilityFunction.CombinedAllSheetsMushakOrBillOfEntryDbDict", "D:\Temp\UP Draft\Draft 2025\Import Performance Statement of PDL-2025-2026.xlsx")
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
    
    ' test1 = Application.Run("general_utility_functions.extractAndFormatUdNo", "BGMEA/DHK/UD/2024/3578/020")
    ' test1 = Application.Run("general_utility_functions.extractAndFormatUdNo", "BGMEA/DHK/AM/2024/3016/002-003")

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
    ' Set dict = Application.Run("dedo_consumption.combineAllDedoConDicAfterCalculateActualQty", 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100)
'
'    test2 = Application.Run("dictionary_utility_functions.sumOfProvidedKeys", dict, test1)
'
'
'    Set dict = Application.Run("dedo_consumption.appliedNotUsedRawMaterials", dict, test1)
'
'    test2 = Application.Run("dictionary_utility_functions.sumOfProvidedKeys", dict, test1)
'
    ' Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", Range("a1"), dict, 1, 1, 1

    ' Set dict = Application.Run("dedo_consumption.finalRawMaterialsQtyCalculatedAsGroup", 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100)

    
'    Application.Run "afterConsumption.upClause8MakeUniqueRowsFromProvidedWs", ActiveSheet

'    test2 = Application.Run("general_utility_functions.ExtractFirstLineWithRegex", "sampleStr654")

'    test2 = Application.Run("general_utility_functions.dictKeyGeneratorWithLcMushakOrBillOfEntryQtyAndValue", "654654-L", "c-654", 654.87, 598.657)

        ' Set test2 = Application.Run("general_utility_functions.upClause8InformationFromProvidedWs", ActiveSheet)

        ' Set test2 = Application.Run("general_utility_functions.sumUsedQtyAndValueAsMushakOrBillOfEntryFromSelectedUpFile")

        ' Set dict = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", "D:\Temp\UP Draft\Draft 2025\json-used-up-clause8\file.json")
        ' Set dict = Application.Run("Sorting_Algorithms.SplituPSequence", Array("2/2024", "3/2024", "5/2024"))
        
        ' test2 = Application.Run("utilityFunction.upSequenceStrGenerator", Array("2/2024", "3/2024", "5/2024"))

        Application.Run "yarnConsumption.yarnConsumptionInformationPutToProvidedWs", ActiveSheet.Range("a5:aa6"), 1, CreateObject("Scripting.Dictionary")
    
    End Sub
    

