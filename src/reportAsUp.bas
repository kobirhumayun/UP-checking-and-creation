Attribute VB_Name = "reportAsUp"
Option Explicit

Private Function copySmpleFileAsNewReportFileAndReturnAllPath(basePath As String, sampleUpFilePathDeem As String, sampleUpFilePathDirect As String, totalUpListForReport As Variant, allUpDicFromJson As Object) As Object

    Dim deemUpFullPathDict As Object
    Set deemUpFullPathDict = CreateObject("Scripting.Dictionary")
    
    Dim directUpFullPathDict As Object
    Set directUpFullPathDict = CreateObject("Scripting.Dictionary")
    
    Dim upNotFoundInAllUpDicFromJson As Object
    Set upNotFoundInAllUpDicFromJson = CreateObject("Scripting.Dictionary")
    
    Dim element As Variant
    
        'create all file path for report
    For Each element In totalUpListForReport
    
        If allUpDicFromJson.Exists(element) Then
        
            If allUpDicFromJson(element)("upClause7")("1")("isGarments") Or allUpDicFromJson(element)("upClause7")("1")("isExistIp") Or allUpDicFromJson(element)("upClause7")("1")("isExistExp") Then
            
                    'direct UP path
                directUpFullPathDict.Add element, basePath & Application.PathSeparator & "UP-" & Replace(element, "/", "-") & "-Import-Export-UP-Performance-Direct.xlsx"
                
            Else
            
                    'deem UP path
                deemUpFullPathDict.Add element, basePath & Application.PathSeparator & "UP-" & Replace(element, "/", "-") & "-Import-Export-UP-Performance-Deem.xlsx"
             
            End If
            
        Else
                'UP not found in json data
            upNotFoundInAllUpDicFromJson.Add element, element
            
        End If
        
    Next element

    Dim uPSequenceStr As String
    
        'if source data not found show msg. & stop process
    If upNotFoundInAllUpDicFromJson.Count > 0 Then
    
        uPSequenceStr = Application.Run("utilityFunction.upSequenceStrGenerator", upNotFoundInAllUpDicFromJson.keys, " -to- ", 10)
        
        MsgBox "UP not found in source data" & Chr(10) & "Generate JSON Dictionary first" & Chr(10) & uPSequenceStr
        Exit Function
        
    End If
    
    Dim outerKey As Variant
    
    Dim previouslyReportFileWasCreated As Object
    Set previouslyReportFileWasCreated = CreateObject("Scripting.Dictionary")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
        
        'Remove previously created deem report path & keep record
    For Each outerKey In deemUpFullPathDict.keys
    
        If fso.FileExists(deemUpFullPathDict(outerKey)) Then
        
            previouslyReportFileWasCreated.Add outerKey, outerKey
            deemUpFullPathDict.Remove outerKey
    
        End If

    Next outerKey
    
        'Remove previously created direct report path & keep record
    For Each outerKey In directUpFullPathDict.keys
    
        If fso.FileExists(directUpFullPathDict(outerKey)) Then
            
            previouslyReportFileWasCreated.Add outerKey, outerKey
            directUpFullPathDict.Remove outerKey
    
        End If

    Next outerKey
    
        'if previously created report exist show msg. for awareness
    If previouslyReportFileWasCreated.Count > 0 Then
    
        uPSequenceStr = Application.Run("utilityFunction.upSequenceStrGenerator", previouslyReportFileWasCreated.keys, " -to- ", 10)
        
        MsgBox "UP report previously created" & Chr(10) & "Skip these UP" & Chr(10) & uPSequenceStr
        
    End If
    
        'copy deem sample file as new report file
    For Each outerKey In deemUpFullPathDict.keys
    
        Application.Run "general_utility_functions.CopyFileAsNewFileFSO", sampleUpFilePathDeem, deemUpFullPathDict(outerKey), False

    Next outerKey
    
        'copy direct sample file as new report file
    For Each outerKey In directUpFullPathDict.keys
    
        Application.Run "general_utility_functions.CopyFileAsNewFileFSO", sampleUpFilePathDirect, directUpFullPathDict(outerKey), False

    Next outerKey
    
    Dim returnDict As Object
    Set returnDict = CreateObject("Scripting.Dictionary")
    
    returnDict.Add "deemUpFullPathDict", deemUpFullPathDict
    returnDict.Add "directUpFullPathDict", directUpFullPathDict
    
    Set copySmpleFileAsNewReportFileAndReturnAllPath = returnDict
    
End Function

Private Function putValueToReportDeemUp(allUpDicFromJson As Object, deemUpFullPathDict As Object, upNoAndDtAsDict As Object)

    Dim currentReportWb As Workbook
    Dim currentReportWs As Worksheet
    Dim currentReportRange As Range
    
    Dim outerKey As Variant
    Dim rowTracker As Long
    
    Application.ScreenUpdating = False

    For Each outerKey In deemUpFullPathDict.keys
        
        Set currentReportWb = Workbooks.Open(deemUpFullPathDict(outerKey))
        Set currentReportWs = currentReportWb.Worksheets(1)
        Set currentReportRange = currentReportWs.Range("A6:Q20")
        
        With currentReportRange
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Application.Run "reportAsUp.putValueToReportUpColumn", currentReportRange.Columns("a"), allUpDicFromJson(outerKey)("upClause1"), upNoAndDtAsDict(outerKey)("upDt")
        
        Application.Run "reportAsUp.putValueToReportExportLcColumn", currentReportRange.Columns("b"), allUpDicFromJson(outerKey)("upClause7")
        
        Application.Run "reportAsUp.putValueToReportRawMaterialsQtyColumn", currentReportRange.Columns("c"), allUpDicFromJson(outerKey)("upClause13")

        Dim divideIntoImportAndLocalLc As Object
        Set divideIntoImportAndLocalLc = Application.Run("reportAsUp.divideIntoImportAndLocalLc", allUpDicFromJson(outerKey)("upClause8"))

        Dim groupByLcAndRawMaterialsLocal As Object
        Set groupByLcAndRawMaterialsLocal = Application.Run("reportAsUp.groupByLcAndRawMaterials", divideIntoImportAndLocalLc("localLc"))

        Dim groupByLcAndRawMaterialsImport As Object
        Set groupByLcAndRawMaterialsImport = Application.Run("reportAsUp.groupByLcAndRawMaterials", divideIntoImportAndLocalLc("importLc"))

        Application.Run "reportAsUp.putValueToReportLcValueQtyColumn", currentReportRange.Columns("d:g"), groupByLcAndRawMaterialsImport
        Application.Run "reportAsUp.putValueToReportLcValueQtyColumn", currentReportRange.Columns("h:k"), groupByLcAndRawMaterialsLocal
        
        Application.Run "reportAsUp.putValueToReportFabricsQtyColumn", currentReportRange.Columns("l"), allUpDicFromJson(outerKey)("upClause7")
        Application.Run "reportAsUp.putValueToReportFabricsQtyColumn", currentReportRange.Columns("m"), allUpDicFromJson(outerKey)("upClause7")
        
        Application.Run "reportAsUp.putValueToReportBuyerNameColumn", currentReportRange.Columns("n"), allUpDicFromJson(outerKey)("upClause6")

        Application.Run "reportAsUp.putValueToReportExportValueColumn", currentReportRange.Columns("o"), allUpDicFromJson(outerKey)("upClause7")

        currentReportWb.Close SaveChanges:=True
    
    Next outerKey
    
    Application.ScreenUpdating = True

End Function

Private Function putValueToReportDirectUp(allUpDicFromJson As Object, directUpFullPathDict As Object, upNoAndDtAsDict As Object)

    Dim currentReportWb As Workbook
    Dim currentReportWs As Worksheet
    Dim currentReportRange As Range
    
    Dim outerKey As Variant
    Dim rowTracker As Long
    
    Application.ScreenUpdating = False

    For Each outerKey In directUpFullPathDict.keys
        
        Set currentReportWb = Workbooks.Open(directUpFullPathDict(outerKey))
        Set currentReportWs = currentReportWb.Worksheets(1)
        Set currentReportRange = currentReportWs.Range("A6:Q20")
        
        With currentReportRange
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Application.Run "reportAsUp.putValueToReportExportLcColumn", currentReportRange.Columns("a"), allUpDicFromJson(outerKey)("upClause7")
        
        Application.Run "reportAsUp.putValueToReportUpColumn", currentReportRange.Columns("b"), allUpDicFromJson(outerKey)("upClause1"), upNoAndDtAsDict(outerKey)("upDt")
        
        Application.Run "reportAsUp.putValueToReportRawMaterialsQtyColumn", currentReportRange.Columns("c"), allUpDicFromJson(outerKey)("upClause13")

        Dim divideIntoImportAndLocalLc As Object
        Set divideIntoImportAndLocalLc = Application.Run("reportAsUp.divideIntoImportAndLocalLc", allUpDicFromJson(outerKey)("upClause8"))

        Dim groupByLcAndRawMaterialsLocal As Object
        Set groupByLcAndRawMaterialsLocal = Application.Run("reportAsUp.groupByLcAndRawMaterials", divideIntoImportAndLocalLc("localLc"))

        Dim groupByLcAndRawMaterialsImport As Object
        Set groupByLcAndRawMaterialsImport = Application.Run("reportAsUp.groupByLcAndRawMaterials", divideIntoImportAndLocalLc("importLc"))

        Application.Run "reportAsUp.putValueToReportLcValueQtyColumn", currentReportRange.Columns("d:g"), groupByLcAndRawMaterialsImport
        Application.Run "reportAsUp.putValueToReportLcValueQtyColumn", currentReportRange.Columns("h:k"), groupByLcAndRawMaterialsLocal

        Application.Run "reportAsUp.putValueToReportExportValueColumn", currentReportRange.Columns("o"), allUpDicFromJson(outerKey)("upClause7")

        currentReportWb.Close SaveChanges:=True
    
    Next outerKey
    
    Application.ScreenUpdating = True

End Function

Private Function putValueToReportUpColumn(upRange As Range, upClause1 As Object, upDate As Date)

    Dim rowTracker As Long
    rowTracker = 1
    
    upRange.Range("a" & rowTracker).NumberFormat = "@"
    upRange.Range("a" & rowTracker).value = upClause1("upNo")
    
    rowTracker = rowTracker + 1
    upRange.Range("a" & rowTracker).NumberFormat = "dd/mm/yyyy"
    upRange.Range("a" & rowTracker).value = CDate(upDate)
        
End Function

Private Function putValueToReportExportLcColumn(exportLcRange As Range, upClause7 As Object)

    Dim rowTracker As Long
    rowTracker = 1
        
    Dim outerKey As Variant
        
    For Each outerKey In upClause7.keys
        
            'skip first row, then each iteration move two row, extra one row for blank
        If rowTracker > 1 Then
            rowTracker = rowTracker + 2
        End If
        
            'insert two or one rows, due to rowTracker move two rows down
        If ((exportLcRange.Rows.Count - rowTracker) <= 3) Then
                'insert one or two rows
            If ((exportLcRange.Rows.Count - rowTracker) = 3) Then
                    'insert one row, if rowTracker point second from the end
                    'insert above last two rows, to keep format according
                exportLcRange.Rows(exportLcRange.Rows.Count - 1).EntireRow.Insert

            Else
                    'insert two rows, if rowTracker point last row
                    'insert above last two rows, to keep format according
                exportLcRange.Rows(exportLcRange.Rows.Count - 1).EntireRow.Insert
                exportLcRange.Rows(exportLcRange.Rows.Count - 1).EntireRow.Insert
            End If
            
        End If
        
        exportLcRange.Range("a" & rowTracker).NumberFormat = "@"
        exportLcRange.Range("a" & rowTracker).value = upClause7(outerKey)("lcNo")
        
        rowTracker = rowTracker + 1
                'insert one row, if rowTracker point second from the end
            If ((exportLcRange.Rows.Count - rowTracker) = 3) Then
                    'insert above last two rows, to keep format according
                exportLcRange.Rows(exportLcRange.Rows.Count - 1).EntireRow.Insert

            End If
        
        exportLcRange.Range("a" & rowTracker).NumberFormat = "dd/mm/yyyy"
        exportLcRange.Range("a" & rowTracker).value = CDate(upClause7(outerKey)("lcDt"))
        
        If upClause7(outerKey)("isLcAmndExist") Then
            
            rowTracker = rowTracker + 1
            
                'insert one row, if rowTracker point second from the end
            If ((exportLcRange.Rows.Count - rowTracker) = 3) Then
                    'insert above last two rows, to keep format according
                exportLcRange.Rows(exportLcRange.Rows.Count - 1).EntireRow.Insert

            End If
        
            exportLcRange.Range("a" & rowTracker).NumberFormat = "@"
            
            If upClause7(outerKey)("lcAmndNo") < 10 Then
                exportLcRange.Range("a" & rowTracker).value = "Amnd-0" & upClause7(outerKey)("lcAmndNo")
            Else
                exportLcRange.Range("a" & rowTracker).value = "Amnd-" & upClause7(outerKey)("lcAmndNo")
            End If
            
            rowTracker = rowTracker + 1
            
                'insert one row, if rowTracker point second from the end
            If ((exportLcRange.Rows.Count - rowTracker) = 3) Then
                    'insert above last two rows, to keep format according
                exportLcRange.Rows(exportLcRange.Rows.Count - 1).EntireRow.Insert

            End If
            
            exportLcRange.Range("a" & rowTracker).NumberFormat = "dd/mm/yyyy"
            exportLcRange.Range("a" & rowTracker).value = CDate(upClause7(outerKey)("lcAmndDt"))
            
        End If
        
    Next outerKey
      
    exportLcRange.Rows(exportLcRange.Rows.Count - 1).EntireRow.Insert
    exportLcRange.Rows(exportLcRange.Rows.Count - 1).EntireRow.Insert
    exportLcRange.Rows(exportLcRange.Rows.Count - 1).EntireRow.Insert
    
    Dim exportQty, exportValue As Variant
    
    exportQty = Application.Run("dictionary_utility_functions.sumOfInnerDictOfProvidedKeys", upClause7, Array("fabricsQtyInYds"))
    exportValue = Application.Run("dictionary_utility_functions.sumOfInnerDictOfProvidedKeys", upClause7, Array("lcValueInUsd"))
    
    rowTracker = rowTracker + 2
    exportLcRange.Range("a" & rowTracker).NumberFormat = "#,##0.00 ""YDS"""
    exportLcRange.Range("a" & rowTracker).value = exportQty
    
    rowTracker = rowTracker + 1
    exportLcRange.Range("a" & rowTracker).NumberFormat = "$#,##0.00"
    exportLcRange.Range("a" & rowTracker).value = exportValue
    
End Function

Private Function putValueToReportRawMaterialsQtyColumn(upRange As Range, upClause13 As Object)

    Dim rowTracker As Long
    rowTracker = 1
    
    upRange.Range("a" & rowTracker).NumberFormat = "@"
    upRange.Range("a" & rowTracker).value = "Yarn: " & WorksheetFunction.Text(upClause13("yarnImport")("qty") + upClause13("yarnLocal")("qty"), "#,##0.00") & " Kgs"
    
    rowTracker = rowTracker + 1
    upRange.Range("a" & rowTracker).NumberFormat = "@"
    upRange.Range("a" & rowTracker).value = "Dyes: " & WorksheetFunction.Text(upClause13("dyes")("qty"), "#,##0.00") & " Kgs"

    rowTracker = rowTracker + 1
    upRange.Range("a" & rowTracker).NumberFormat = "@"
    upRange.Range("a" & rowTracker).value = "Chem.: " & WorksheetFunction.Text(upClause13("chemicalsImport")("qty") + upClause13("chemicalsLocal")("qty"), "#,##0.00") & " Kgs"

    rowTracker = rowTracker + 1
    upRange.Range("a" & rowTracker).NumberFormat = "@"
    upRange.Range("a" & rowTracker).value = "St. Flim : " & WorksheetFunction.Text(upClause13("stretchWrappingFilm")("qty"), "#,##0.00") & " Kgs"

End Function

Private Function divideIntoImportAndLocalLc(upClause8 As Object) As Object

    Dim bothImportAndLocalLc As Object
    Set bothImportAndLocalLc = CreateObject("Scripting.Dictionary")
    
    Dim importLc As Object
    Set importLc = CreateObject("Scripting.Dictionary")
    
    Dim localLc As Object
    Set localLc = CreateObject("Scripting.Dictionary")
    
    Dim outerKey As Variant
        
    For Each outerKey In upClause8.keys
    
        If Application.Run("general_utility_functions.isStrPatternExist", upClause8(outerKey)("mushakOrBillOfEntryNoAndDt"), "^c-", True, True, True) Then
            importLc.Add outerKey, upClause8(outerKey)
        Else
            localLc.Add outerKey, upClause8(outerKey)
        End If
        
    Next outerKey
    
    bothImportAndLocalLc.Add "importLc", importLc
    bothImportAndLocalLc.Add "localLc", localLc
    
    Set divideIntoImportAndLocalLc = bothImportAndLocalLc
    
End Function

Private Function groupByLcAndRawMaterials(upClause8 As Object) As Object
    
    Dim tempGroup As Object
    Set tempGroup = CreateObject("Scripting.Dictionary")
    
    Dim outerKey As Variant
    Dim tempDictKey As String
    Dim lcNo As String
    
    For Each outerKey In upClause8.keys
    
        lcNo = Application.Run("general_utility_functions.ExtractFirstLineWithRegex", upClause8(outerKey)("lcNoAndDt"))
        lcNo = Replace(lcNo, Chr(13), "")

        tempDictKey = Application.Run("general_utility_functions.RemoveInvalidChars", lcNo) & "_" & Application.Run("general_utility_functions.RemoveInvalidChars", upClause8(outerKey)("nameOfGoods"))
            
        If Not tempGroup.Exists(tempDictKey) Then
        
            tempGroup.Add tempDictKey, CreateObject("Scripting.Dictionary")
            tempGroup(tempDictKey)("lcNo") = lcNo
            tempGroup(tempDictKey)("lcDt") = Right(upClause8(outerKey)("lcNoAndDt"), 10)
            tempGroup(tempDictKey)("nameOfGoods") = upClause8(outerKey)("nameOfGoods")
            tempGroup(tempDictKey)("qty") = 0
            tempGroup(tempDictKey)("value") = 0
            
        End If
        
        tempGroup(tempDictKey)("qty") = tempGroup(tempDictKey)("qty") + upClause8(outerKey)("inThisUpUsedQtyOfGoods")
        tempGroup(tempDictKey)("value") = tempGroup(tempDictKey)("value") + upClause8(outerKey)("inThisUpUsedValueOfGoods")
    
    Next outerKey
    
    Set groupByLcAndRawMaterials = tempGroup
    
End Function

Private Function putValueToReportLcValueQtyColumn(lcRange As Range, groupByLc As Object)

    Dim rowTracker As Long
    rowTracker = 1
        
    Dim outerKey As Variant
    
    For Each outerKey In groupByLc.keys
        
        If groupByLc(outerKey)("qty") = 0 Then
            groupByLc.Remove outerKey
        End If
    
    Next outerKey
    
    For Each outerKey In groupByLc.keys
        
            'skip first row, then each iteration move two row, extra one row for blank
        If rowTracker > 1 Then
            rowTracker = rowTracker + 2
        End If
        
            'insert two or one rows, due to rowTracker move two rows down
        If ((lcRange.Rows.Count - rowTracker) <= 3) Then
                'insert one or two rows
            If ((lcRange.Rows.Count - rowTracker) = 3) Then
                    'insert one row, if rowTracker point second from the end
                    'insert above last two rows, to keep format according
                lcRange.Rows(lcRange.Rows.Count - 1).EntireRow.Insert

            Else
                    'insert two rows, if rowTracker point last row
                    'insert above last two rows, to keep format according
                lcRange.Rows(lcRange.Rows.Count - 1).EntireRow.Insert
                lcRange.Rows(lcRange.Rows.Count - 1).EntireRow.Insert
            End If
            
        End If
        
        lcRange.Range("a" & rowTracker).NumberFormat = "@"
        lcRange.Range("a" & rowTracker).value = groupByLc(outerKey)("lcNo")
        
        With lcRange.Range("b" & rowTracker).Resize(2)
            .VerticalAlignment = xlTop
            .NumberFormat = "@"
            .WrapText = True
            .MergeCells = True
        End With
        lcRange.Range("b" & rowTracker).value = groupByLc(outerKey)("nameOfGoods")
        
        lcRange.Range("c" & rowTracker).NumberFormat = "$#,##0.00"
        lcRange.Range("c" & rowTracker).value = groupByLc(outerKey)("value")
        
        lcRange.Range("d" & rowTracker).Style = "Comma"
        lcRange.Range("d" & rowTracker).value = groupByLc(outerKey)("qty")
        
        rowTracker = rowTracker + 1
                'insert one row, if rowTracker point second from the end
            If ((lcRange.Rows.Count - rowTracker) = 3) Then
                    'insert above last two rows, to keep format according
                lcRange.Rows(lcRange.Rows.Count - 1).EntireRow.Insert

            End If
        
        lcRange.Range("a" & rowTracker).NumberFormat = "dd/mm/yyyy"
        lcRange.Range("a" & rowTracker).value = CDate(groupByLc(outerKey)("lcDt"))
        
    Next outerKey
        
End Function

Private Function putValueToReportBuyerNameColumn(buyerNameRange As Range, upClause6 As Object)

    Dim rowTracker As Long
    rowTracker = 1
    
    Dim loopCounter As Long
    loopCounter = 0
    
    Dim outerKey As Variant
        
    For Each outerKey In upClause6.keys
        
        loopCounter = loopCounter + 1
        
        buyerNameRange.Range("a" & rowTracker).value = upClause6(outerKey)
        
        With buyerNameRange.Range("a" & rowTracker).Resize(3)
            .VerticalAlignment = xlTop
            .NumberFormat = "@"
            .WrapText = True
            .Merge
        End With

        rowTracker = rowTracker + 4
        
        If loopCounter < upClause6.Count Then
                'insert two or one rows, due to rowTracker move two rows down
            If ((buyerNameRange.Rows.Count - rowTracker) <= 3) Then
                    'insert one or two rows
                If ((buyerNameRange.Rows.Count - rowTracker) = 3) Then
                        'insert one row, if rowTracker point second from the end
                        'insert above last two rows, to keep format according
                    buyerNameRange.Rows(buyerNameRange.Rows.Count - 1).EntireRow.Insert
                    buyerNameRange.Rows(buyerNameRange.Rows.Count - 1).EntireRow.Insert
                    buyerNameRange.Rows(buyerNameRange.Rows.Count - 1).EntireRow.Insert
    
                Else
                        'insert two rows, if rowTracker point last row
                        'insert above last two rows, to keep format according
                    buyerNameRange.Rows(buyerNameRange.Rows.Count - 1).EntireRow.Insert
                    buyerNameRange.Rows(buyerNameRange.Rows.Count - 1).EntireRow.Insert
                    buyerNameRange.Rows(buyerNameRange.Rows.Count - 1).EntireRow.Insert
                    buyerNameRange.Rows(buyerNameRange.Rows.Count - 1).EntireRow.Insert
                End If
                
            End If
        
        End If

    Next outerKey
        
End Function

Private Function putValueToReportFabricsQtyColumn(fabricsQtyRange As Range, upClause7 As Object)

    Dim rowTracker As Long
    rowTracker = 1
        
    Dim fabricsQty As Variant
    fabricsQty = 0

    Dim outerKey As Variant
        
    For Each outerKey In upClause7.keys
        
        fabricsQty = fabricsQty + upClause7(outerKey)("fabricsQtyInYds")
        
    Next outerKey

    fabricsQtyRange.Range("a" & rowTracker).Style = "Comma"
    fabricsQtyRange.Range("a" & rowTracker).value = fabricsQty
        
End Function

Private Function putValueToReportExportValueColumn(exportValueRange As Range, upClause7 As Object)

    Dim rowTracker As Long
    rowTracker = 1
        
    Dim exportValue As Variant
    exportValue = 0

    Dim outerKey As Variant
        
    For Each outerKey In upClause7.keys
        
        exportValue = exportValue + upClause7(outerKey)("lcValueInUsd")
        
    Next outerKey

    exportValueRange.Range("a" & rowTracker).NumberFormat = "$#,##0.00"
    exportValueRange.Range("a" & rowTracker).value = exportValue
        
End Function

Private Function upNoAndDtAsDict(upNoAndDtFilePath As String) As Object

    Application.ScreenUpdating = False

    Dim upNoAndDtWb As Workbook
    Dim upNoAndDtWs As Worksheet
    Dim upNoAndDtRange As Range
    
    Set upNoAndDtWb = Workbooks.Open(upNoAndDtFilePath)
    Set upNoAndDtWs = upNoAndDtWb.Worksheets(1)
    Set upNoAndDtRange = upNoAndDtWs.Range("A1:B" & upNoAndDtWs.Range("A1").End(xlDown).Row)
    
    Dim rangeDataArr As Variant
    rangeDataArr = upNoAndDtRange.value
    
    upNoAndDtWb.Close SaveChanges:=False

    Application.ScreenUpdating = True

    Dim upNoAndDtDict As Object
    Set upNoAndDtDict = CreateObject("Scripting.Dictionary")
    
    Dim tempDictKey As Variant
    
    Dim i As Long
    
    For i = LBound(rangeDataArr) To UBound(rangeDataArr)
        
        tempDictKey = Trim(rangeDataArr(i, 1))
        
        If Not upNoAndDtDict.Exists(tempDictKey) Then
        
            upNoAndDtDict.Add tempDictKey, CreateObject("Scripting.Dictionary")
            upNoAndDtDict(tempDictKey)("upNo") = rangeDataArr(i, 1)
            upNoAndDtDict(tempDictKey)("upDt") = rangeDataArr(i, 2)
        
        End If
        
    Next i
    
    Set upNoAndDtAsDict = upNoAndDtDict
      
End Function

Private Function isExistRelatedUpDate(upNoAndDtAsDict As Object, totalUpListForReport As Variant) As Boolean

    Dim upNotFoundInTotalUpListForReport As Object
    Set upNotFoundInTotalUpListForReport = CreateObject("Scripting.Dictionary")
    
    Dim element As Variant
    
    For Each element In totalUpListForReport
    
        If Not upNoAndDtAsDict.Exists(element) Then
        
                'UP not found in totalUpListForReport
            upNotFoundInTotalUpListForReport.Add element, element
        
        End If
    
    Next element
    
    Dim uPSequenceStr As String
    
        'if source data not found show msg. & stop process
    If upNotFoundInTotalUpListForReport.Count > 0 Then
    
        uPSequenceStr = Application.Run("utilityFunction.upSequenceStrGenerator", upNotFoundInTotalUpListForReport.keys, " -to- ", 10)
        
        MsgBox "UP Date not found in source data" & Chr(10) & "Ensure below UP date in source data" & Chr(10) & uPSequenceStr
        
        isExistRelatedUpDate = False
        Exit Function
    
    End If
    
    isExistRelatedUpDate = True
    
End Function
