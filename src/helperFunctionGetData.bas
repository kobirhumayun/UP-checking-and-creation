Attribute VB_Name = "helperFunctionGetData"
Option Explicit

Private Function upNo() As Variant
'this function give the UP no. from active sheet
Dim temp As Variant
Dim regex As New RegExp
regex.Global = True
regex.MultiLine = True
Range("N13").Interior.ColorIndex = 23
temp = ActiveSheet.Range("N13").value

regex.pattern = "\d+\/\d+"
Set temp = regex.Execute(temp)

upNo = temp.Item(0)

End Function

Private Function upNoFromProvidedWs(ws As Worksheet) As Variant
'this function give the UP no. from provided sheet
Dim temp As Variant
Dim regex As New RegExp
regex.Global = True
regex.MultiLine = True
ws.Range("N13").Interior.ColorIndex = 23
temp = ws.Range("N13").value

regex.pattern = "\d+\/\d+"
Set temp = regex.Execute(temp)

upNoFromProvidedWs = temp.Item(0)

End Function

Private Function upClause6Buyerinformation() As Variant
'this function give buyer information from active sheet

Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("6|", LookAt:=xlWhole).Row
bottomRow = Range("N" & Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("localB2bLcDesBengaliTxt"), LookAt:=xlPart).Row).End(xlUp).Row

Dim workingRange As Range
Set workingRange = Range("N" & topRow & ":" & "N" & bottomRow)
workingRange.Interior.ColorIndex = 23
upClause6Buyerinformation = workingRange.value

End Function

Private Function upClause7Lcinformation() As Variant
'this function give LC information from active sheet

Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("localB2bLcDesBengaliTxt"), LookAt:=xlPart).Row + 2
bottomRow = Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row - 1

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AA" & bottomRow)
workingRange.Interior.ColorIndex = 23
upClause7Lcinformation = workingRange.value

End Function

Private Function upClause8BtbLcinformation() As Variant
'this function give BTB LC information from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row + 3
bottomRow = Range("V" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AA" & bottomRow)
workingRange.Interior.ColorIndex = 23
upClause8BtbLcinformation = workingRange.value

End Function

Private Function upClause8BtbLcinformationPreviousFormat() As Variant
    'this function give BTB LC information from active sheet
    Dim topRow, bottomRow As Variant

    topRow = ActiveSheet.Cells.Find("8|  Avg`vbx Gj/wm Gi weeiY", LookAt:=xlPart).Row + 2
    bottomRow = Range("S" & topRow).End(xlDown).Row

    Dim workingRange As Range
    Set workingRange = Range("B" & topRow & ":" & "X" & bottomRow)
    workingRange.Interior.ColorIndex = 23
    upClause8BtbLcinformationPreviousFormat = workingRange.value

End Function

Private Function upClause9Stockinformation() As Variant
'this function give stock information from active sheet

Dim vsCodeNotSupportedOrBengaliTxtDictionary As Object
Set vsCodeNotSupportedOrBengaliTxtDictionary = Application.Run("vs_code_not_supported_text.CreateVsCodeNotSupportedOrBengaliTxtDictionary")

Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find(vsCodeNotSupportedOrBengaliTxtDictionary("infoAboutStockBengaliTxt"), LookAt:=xlPart).Row + 3
bottomRow = Range("T" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AC" & bottomRow)
workingRange.Interior.ColorIndex = 23
upClause9Stockinformation = workingRange.value

End Function

Private Function upClause11UdExpIpinformation() As Variant
'this function give UD/EXP/IP information from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("11|", LookAt:=xlPart).Row + 3
bottomRow = Range("Z" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AA" & bottomRow)
workingRange.Interior.ColorIndex = 23
upClause11UdExpIpinformation = workingRange.value

End Function

Private Function upClause12AYarnConsumptioninformation() As Variant
'this function give yarn consumption information from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("12| (K)", LookAt:=xlPart).Row + 2
bottomRow = Range("Z" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "AA" & bottomRow)
workingRange.Interior.ColorIndex = 23
upClause12AYarnConsumptioninformation = workingRange.value

End Function

Private Function upClause12BChemicalDyesConsumptioninformation() As Variant
'this function give chemical & dyes consumption information from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("12| (L)", LookAt:=xlPart).Row + 2
bottomRow = Range("X" & topRow + 1).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "Y" & bottomRow)
workingRange.Interior.ColorIndex = 23
upClause12BChemicalDyesConsumptioninformation = workingRange.value

End Function

Private Function upClause13UseRawMaterialsinformation() As Variant
'this function give used raw materials information from active sheet
Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("13|", LookAt:=xlPart).Row + 2
bottomRow = Range("R" & topRow).End(xlDown).Row

Dim workingRange As Range
Set workingRange = Range("B" & topRow & ":" & "R" & bottomRow)
workingRange.Interior.ColorIndex = 23
upClause13UseRawMaterialsinformation = workingRange.value


End Function


Private Function upYarnConsumptionInformation() As Variant
'this function give yarn consumption information from consumption sheet

Dim upSheetName As Variant

upSheetName = ActiveSheet.Name
ActiveWorkbook.Worksheets("Consumption").Activate


Dim topRow, bottomRow As Variant

topRow = ActiveSheet.Cells.Find("TOTAL", LookAt:=xlPart).Row
bottomRow = topRow + 12

Dim workingRange As Range
Set workingRange = Range("C" & topRow & ":" & "N" & bottomRow)
workingRange.Interior.ColorIndex = 23

ActiveWorkbook.Worksheets(upSheetName).Activate

upYarnConsumptionInformation = workingRange.value


End Function



Private Function sourceDataUpIssuingStatus(upNo As String, fileName As String, worksheetTabName As String) As Variant  ' provide UP no., source file name & worksheetTabName
'this function give source data from UP Issuing Status

    Application.Run "utilityFunction.openFile", fileName ' provide filename
    
    ActiveWorkbook.Worksheets(worksheetTabName).Activate
    ActiveSheet.AutoFilterMode = False
    
    Dim workingRange As Range
    Set workingRange = Range("A3:" & "AE" & Range("B2").End(xlDown).Row)
    
    Dim temp As Variant
    temp = workingRange.value
    
    Dim returnArr As Variant
    returnArr = Application.Run("utilityFunction.towDimensionalArrayFilter", temp, "^" & Replace(upNo, "/", "\/") & "$", 24)   ' provide array, pattern string & filter index
    
    Dim i As Integer
    For i = 1 To UBound(returnArr, 1)
        returnArr(i, 27) = Mid(workingRange(returnArr(i, 32), 9).NumberFormat, 11, 3) ' take qty. numberFormat & store column 27 for knowing qty. unit (RangeObject work only when open that file)
    Next i
    
    Application.Run "utilityFunction.closeFile", fileName ' provide filename
    
    sourceDataUpIssuingStatus = returnArr

End Function


Private Function sourceDataImportPerformance(fileName As String, worksheetTabName As String, openFile As Boolean, closeFile As Boolean) As Variant ' provide source file name & worksheetTabName
'this function give source data from Import Performance
    
    
    
    If openFile Then
        Application.Run "utilityFunction.openFile", fileName ' provide filename
    End If
    
    ActiveWorkbook.Worksheets(worksheetTabName).Activate
    ActiveSheet.AutoFilterMode = False
    
    Dim temp As Variant
    temp = Range("A6:" & "N" & Range("C6").End(xlDown).Row).value
    
    
    If closeFile Then
       Application.Run "utilityFunction.closeFile", fileName ' provide filename
    End If
    
    sourceDataImportPerformance = temp

End Function


Private Function sourceDataPreviousUp(currentUp As Variant, upClauseNo As Integer, openFile As Boolean, closeFile As Boolean) As Variant  ' provide UP clause No.
'this function give source data from previous UP
    
    Dim fileName As String
    
    Dim currentUpOnlyNo, currentUpYear, previousUpOnlyNo As Variant
        
    Dim regex As New RegExp
    regex.Global = True
    regex.MultiLine = True

    regex.pattern = "^\d+"
    Set currentUpOnlyNo = regex.Execute(currentUp)
    previousUpOnlyNo = currentUpOnlyNo.Item(0) - 1
    
    regex.pattern = "\d+$"
    Set currentUpYear = regex.Execute(currentUp)
    currentUpYear = currentUpYear.Item(0)
    
    
    
    fileName = "UP-" & previousUpOnlyNo & "-" & currentUpYear & ".xlsx"
    
    
    
    If openFile Then
        Application.Run "utilityFunction.openFile", fileName ' provide filename
    End If
    
    ActiveWorkbook.Worksheets(Worksheets(2).Name).Activate
    ActiveSheet.AutoFilterMode = False
    
    Dim temp As Variant
        
    If upClauseNo = 8 Then
        temp = upClause8BtbLcinformation
    ElseIf upClauseNo = 9 Then
        temp = upClause9Stockinformation
    End If
    
    If closeFile Then
       Application.Run "utilityFunction.closeFile", fileName ' provide filename
    End If
    
    sourceDataPreviousUp = temp

End Function


Private Function sourceDataAsDicUpIssuingStatus(upNo As String, fileName As String, worksheetTabName As String) As Variant  ' provide UP no., source file name & worksheetTabName
        'this function give source data as dictionary from UP Issuing Status
        
        Application.Run "utilityFunction.openFile", fileName ' provide filename
        
        ActiveWorkbook.Worksheets(worksheetTabName).Activate
        ActiveSheet.AutoFilterMode = False
        
        Dim workingRange As Range
        Set workingRange = Range("A2:" & "AH" & Range("B2").End(xlDown).Row)
        
        Dim temp As Variant
        temp = workingRange.value
        
        Dim upIssuingStatusDic As Object
        Set upIssuingStatusDic = CreateObject("Scripting.Dictionary")
        
        Dim lcCount As Object
        Set lcCount = CreateObject("Scripting.Dictionary")
        
        Dim tempLcDic As Object
        
        Dim propertiesArr, propertiesValArr As Variant
        
        ReDim propertiesArr(1 To UBound(temp, 2))
        ReDim propertiesValArr(1 To UBound(temp, 2))
        
        Dim i, j As Long
        
        For j = 1 To UBound(temp, 2)
            propertiesArr(j) = temp(1, j)
        Next j
        
        propertiesArr(6) = "LC" & propertiesArr(6) ' same key conflict handle
        propertiesArr(22) = "BTB" & propertiesArr(22) ' same key conflict handle
        
        For i = 1 To UBound(temp)
          
          If temp(i, 24) = upNo Then
            lcCount(temp(i, 4)) = lcCount(temp(i, 4)) + 1
            
             For j = 1 To UBound(temp, 2)
                 propertiesValArr(j) = temp(i, j)
             Next j
            
            Set tempLcDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)
            
            tempLcDic("currencyNumberFormat") = workingRange(i, 6).NumberFormat
            tempLcDic("qtyNumberFormat") = workingRange(i, 9).NumberFormat
                        
            If Not workingRange(i, 20).Comment Is Nothing Then   'check if the cell has a comment
                tempLcDic("b2bComment") = workingRange(i, 20).Comment.Text
            Else
                tempLcDic("b2bComment") = "No Comment"
            End If
            
                        
            upIssuingStatusDic.Add temp(i, 4) & "_" & lcCount(temp(i, 4)), tempLcDic
           
          End If
          

        Next i

        Application.Run "utilityFunction.closeFile", fileName ' provide filename
                
        Set sourceDataAsDicUpIssuingStatus = Application.Run("dictionary_utility_functions.SortDictionaryByKey", upIssuingStatusDic)

End Function


