Attribute VB_Name = "yarnConsumption"
Option Explicit


Private Function yarnConsumptionInformationPutToProvidedWs(totalConsumptionRange As Range, rowTracker As Long, yarnConsumptionInfoDic As Object)
    'received total range, row No. & yarn info dictionary
    'put yarnConsumption information to provided worksheet & related formula
    '***if any dictionary key and value exit put that right place
    '***if any dictionary key not exit put null value of that place
    'below list of all needed key and value example
    ' yarnConsumptionInfoDic("weight") = 10.75 'should be dynamic
    ' yarnConsumptionInfoDic("width") = 66.5 'should be dynamic
    ' yarnConsumptionInfoDic("fabricQty") = 5000 'should be dynamic
    ' yarnConsumptionInfoDic("black") = "Black" 'add as PI color
    ' yarnConsumptionInfoDic("mercerizationBlack") = "Mercerization(Black)" 'add as PI
    ' yarnConsumptionInfoDic("indigo") = "Indigo" 'add as PI color
    ' yarnConsumptionInfoDic("mercerizationIndigo") = "Mercerization(Indigo)" 'add as PI
    ' yarnConsumptionInfoDic("toppingBottoming") = "Topping/ Bottoming" 'add as PI color
    ' yarnConsumptionInfoDic("mercerizationtoppingBottoming") = "Mercerization(Topping/ Bottoming)" 'add as PI
    ' yarnConsumptionInfoDic("overDying") = "Over Dying" 'add as PI
    ' yarnConsumptionInfoDic("mercerizationoverDying") = "Mercerization(Over Dying)" 'add as PI
    ' yarnConsumptionInfoDic("cottonPercentage") = 90 'should be dynamic
    ' yarnConsumptionInfoDic("coating") = "Coating" 'add as PI
    ' yarnConsumptionInfoDic("polyesterPercentage") = 5 'should be dynamic
    ' yarnConsumptionInfoDic("pfd") = "PFD" 'add as PI
    ' yarnConsumptionInfoDic("spandexPercentage") = 5 'should be dynamic
    ' yarnConsumptionInfoDic("ecru") = "ECRU" 'add as PI

    totalConsumptionRange.Range("a" & rowTracker).value = "Weight :"
    totalConsumptionRange.Range("a" & rowTracker & ":c" & rowTracker).Merge

        'weight
    totalConsumptionRange.Range("d" & rowTracker).value = yarnConsumptionInfoDic("weight")
    totalConsumptionRange.Range("d" & rowTracker).Style = "Comma"
    totalConsumptionRange.Range("d" & rowTracker & ":e" & rowTracker).Merge

    totalConsumptionRange.Range("f" & rowTracker).value = "OZ/YD2"
    totalConsumptionRange.Range("f" & rowTracker & ":g" & rowTracker).Merge

    totalConsumptionRange.Range("i" & rowTracker).value = "Width :"
    totalConsumptionRange.Range("i" & rowTracker & ":k" & rowTracker).Merge

        'Width
    totalConsumptionRange.Range("l" & rowTracker).value = yarnConsumptionInfoDic("width")
    totalConsumptionRange.Range("l" & rowTracker).Style = "Comma"
    totalConsumptionRange.Range("l" & rowTracker & ":n" & rowTracker).Merge

    totalConsumptionRange.Range("o" & rowTracker).value = "Inch"
    totalConsumptionRange.Range("o" & rowTracker & ":p" & rowTracker).Merge

    totalConsumptionRange.Range("r" & rowTracker).value = "Qty :"
    totalConsumptionRange.Range("r" & rowTracker & ":s" & rowTracker).Merge

        'Qty.
    totalConsumptionRange.Range("t" & rowTracker).value = yarnConsumptionInfoDic("fabricQty")
    totalConsumptionRange.Range("t" & rowTracker).Style = "Comma"
    totalConsumptionRange.Range("t" & rowTracker & ":v" & rowTracker).Merge

    totalConsumptionRange.Range("w" & rowTracker).value = "Yds"
    totalConsumptionRange.Range("w" & rowTracker & ":x" & rowTracker).Merge


    totalConsumptionRange.Range("b" & rowTracker + 2).value = "="

        'put formula to take weight
    totalConsumptionRange.Range("c" & rowTracker + 2).FormulaR1C1 = "=R[-2]C[1]"
    totalConsumptionRange.Range("c" & rowTracker + 2).Style = "Comma"
    totalConsumptionRange.Range("c" & rowTracker + 2 & ":d" & rowTracker + 2).Merge

    totalConsumptionRange.Range("e" & rowTracker + 2).value = "x"

        'put formula to take width
    totalConsumptionRange.Range("f" & rowTracker + 2).FormulaR1C1 = "=R[-2]C[6]"
    totalConsumptionRange.Range("f" & rowTracker + 2).Style = "Comma"

    totalConsumptionRange.Range("g" & rowTracker + 2).value = Chr(247)

    totalConsumptionRange.Range("h" & rowTracker + 2).value = 36

    totalConsumptionRange.Range("i" & rowTracker + 2).value = Chr(247)

    totalConsumptionRange.Range("j" & rowTracker + 2).value = 16

    totalConsumptionRange.Range("k" & rowTracker + 2).value = Chr(247)

    totalConsumptionRange.Range("l" & rowTracker + 2).value = 2.2046
    totalConsumptionRange.Range("l" & rowTracker + 2 & ":m" & rowTracker + 2).Merge

    totalConsumptionRange.Range("n" & rowTracker + 2).value = "="

    totalConsumptionRange.Range("o" & rowTracker + 2).FormulaR1C1 = "=RC[-12]*RC[-9]/RC[-7]/RC[-5]/RC[-3]"
    totalConsumptionRange.Range("o" & rowTracker + 2 & ":r" & rowTracker + 2).Merge

    totalConsumptionRange.Range("s" & rowTracker + 2).value = yarnConsumptionInfoDic("black")
    totalConsumptionRange.Range("s" & rowTracker + 2 & ":y" & rowTracker + 2).Merge
    totalConsumptionRange.Range("s" & rowTracker + 2 & ":y" & rowTracker + 2).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 2).value = "Black"
    totalConsumptionRange.Range("ag" & rowTracker + 2 & ":am" & rowTracker + 2).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 2 & ":am" & rowTracker + 2).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("s" & rowTracker + 3).value = yarnConsumptionInfoDic("mercerizationBlack")
    totalConsumptionRange.Range("s" & rowTracker + 3 & ":y" & rowTracker + 3).Merge
    totalConsumptionRange.Range("s" & rowTracker + 3 & ":y" & rowTracker + 3).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 3).value = "Mercerization(Black)"
    totalConsumptionRange.Range("ag" & rowTracker + 3 & ":am" & rowTracker + 3).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 3 & ":am" & rowTracker + 3).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("b" & rowTracker + 4).value = "="

    totalConsumptionRange.Range("c" & rowTracker + 4).FormulaR1C1 = "=R[-2]C[12]"
    totalConsumptionRange.Range("c" & rowTracker + 4 & ":f" & rowTracker + 4).Merge

    totalConsumptionRange.Range("g" & rowTracker + 4).value = "kgs"

    totalConsumptionRange.Range("h" & rowTracker + 4).value = "x"

    totalConsumptionRange.Range("i" & rowTracker + 4).FormulaR1C1 = "=R[-4]C[11]"
    totalConsumptionRange.Range("i" & rowTracker + 4 & ":k" & rowTracker + 4).Merge

    totalConsumptionRange.Range("l" & rowTracker + 4).value = "Yds"
    totalConsumptionRange.Range("l" & rowTracker + 4 & ":m" & rowTracker + 4).Merge

    totalConsumptionRange.Range("s" & rowTracker + 4).value = yarnConsumptionInfoDic("indigo")
    totalConsumptionRange.Range("s" & rowTracker + 4 & ":y" & rowTracker + 4).Merge
    totalConsumptionRange.Range("s" & rowTracker + 4 & ":y" & rowTracker + 4).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 4).value = "Indigo"
    totalConsumptionRange.Range("ag" & rowTracker + 4 & ":am" & rowTracker + 4).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 4 & ":am" & rowTracker + 4).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("s" & rowTracker + 5).value = yarnConsumptionInfoDic("mercerizationIndigo")
    totalConsumptionRange.Range("s" & rowTracker + 5 & ":y" & rowTracker + 5).Merge
    totalConsumptionRange.Range("s" & rowTracker + 5 & ":y" & rowTracker + 5).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 5).value = "Mercerization(Indigo)"
    totalConsumptionRange.Range("ag" & rowTracker + 5 & ":am" & rowTracker + 5).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 5 & ":am" & rowTracker + 5).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("b" & rowTracker + 6).value = "="

    totalConsumptionRange.Range("c" & rowTracker + 6).FormulaR1C1 = "=R[-2]C*R[-2]C[6]"
    totalConsumptionRange.Range("c" & rowTracker + 6).Style = "Comma"
    totalConsumptionRange.Range("c" & rowTracker + 6 & ":f" & rowTracker + 6).Merge

    totalConsumptionRange.Range("g" & rowTracker + 6).value = "kgs"

    totalConsumptionRange.Range("h" & rowTracker + 6).value = "x"

    totalConsumptionRange.Range("i" & rowTracker + 6).value = "6%"
    totalConsumptionRange.Range("i" & rowTracker + 6 & ":j" & rowTracker + 6).Merge

    totalConsumptionRange.Range("s" & rowTracker + 6).value = yarnConsumptionInfoDic("toppingBottoming")
    totalConsumptionRange.Range("s" & rowTracker + 6 & ":y" & rowTracker + 6).Merge
    totalConsumptionRange.Range("s" & rowTracker + 6 & ":y" & rowTracker + 6).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 6).value = "Topping/ Bottoming"
    totalConsumptionRange.Range("ag" & rowTracker + 6 & ":am" & rowTracker + 6).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 6 & ":am" & rowTracker + 6).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("s" & rowTracker + 7).value = yarnConsumptionInfoDic("mercerizationtoppingBottoming")
    totalConsumptionRange.Range("s" & rowTracker + 7 & ":y" & rowTracker + 7).Merge
    totalConsumptionRange.Range("s" & rowTracker + 7 & ":y" & rowTracker + 7).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 7).value = "Mercerization(Topping/ Bottoming)"
    totalConsumptionRange.Range("ag" & rowTracker + 7 & ":am" & rowTracker + 7).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 7 & ":am" & rowTracker + 7).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("b" & rowTracker + 8).value = "="

    totalConsumptionRange.Range("c" & rowTracker + 8).FormulaR1C1 = "=R[-2]C*R[-2]C[6]+R[-2]C"
    totalConsumptionRange.Range("c" & rowTracker + 8).Style = "Comma"
    totalConsumptionRange.Range("c" & rowTracker + 8 & ":f" & rowTracker + 8).Merge

    totalConsumptionRange.Range("g" & rowTracker + 8).value = "kgs"

    totalConsumptionRange.Range("n" & rowTracker + 8).FormulaR1C1 = _
        "=R[-6]C[5]&R[-5]C[5]&R[-4]C[5]&R[-3]C[5]&R[-2]C[5]&R[-1]C[5]&RC[5]&R[1]C[5]&R[2]C[5]&R[3]C[5]&R[4]C[5]"

    totalConsumptionRange.Range("n" & rowTracker + 8).NumberFormat = ";;;" 'hide text

    totalConsumptionRange.Range("s" & rowTracker + 8).value = yarnConsumptionInfoDic("overDying")
    totalConsumptionRange.Range("s" & rowTracker + 8 & ":y" & rowTracker + 8).Merge
    totalConsumptionRange.Range("s" & rowTracker + 8 & ":y" & rowTracker + 8).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 8).value = "Over Dying"
    totalConsumptionRange.Range("ag" & rowTracker + 8 & ":am" & rowTracker + 8).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 8 & ":am" & rowTracker + 8).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("s" & rowTracker + 9).value = yarnConsumptionInfoDic("mercerizationoverDying")
    totalConsumptionRange.Range("s" & rowTracker + 9 & ":y" & rowTracker + 9).Merge
    totalConsumptionRange.Range("s" & rowTracker + 9 & ":y" & rowTracker + 9).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 9).value = "Mercerization(Over Dying)"
    totalConsumptionRange.Range("ag" & rowTracker + 9 & ":am" & rowTracker + 9).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 9 & ":am" & rowTracker + 9).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("a" & rowTracker + 10).value = "Cotton"
    totalConsumptionRange.Range("a" & rowTracker + 10 & ":d" & rowTracker + 10).Merge
    totalConsumptionRange.Range("a" & rowTracker + 10 & ":d" & rowTracker + 10).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("e" & rowTracker + 10).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("f" & rowTracker + 10).FormulaR1C1 = "=R[-2]C[-3]*" & yarnConsumptionInfoDic("cottonPercentage") & "%"
    totalConsumptionRange.Range("f" & rowTracker + 10).Style = "Comma"
    totalConsumptionRange.Range("f" & rowTracker + 10 & ":j" & rowTracker + 10).Merge
    totalConsumptionRange.Range("f" & rowTracker + 10 & ":j" & rowTracker + 10).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("s" & rowTracker + 10).value = yarnConsumptionInfoDic("coating")
    totalConsumptionRange.Range("s" & rowTracker + 10 & ":y" & rowTracker + 10).Merge
    totalConsumptionRange.Range("s" & rowTracker + 10 & ":y" & rowTracker + 10).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 10).value = "Coating"
    totalConsumptionRange.Range("ag" & rowTracker + 10 & ":am" & rowTracker + 10).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 10 & ":am" & rowTracker + 10).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("a" & rowTracker + 11).value = "Polyester"
    totalConsumptionRange.Range("a" & rowTracker + 11 & ":d" & rowTracker + 11).Merge
    totalConsumptionRange.Range("a" & rowTracker + 11 & ":d" & rowTracker + 11).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("e" & rowTracker + 11).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("f" & rowTracker + 11).FormulaR1C1 = "=R[-3]C[-3]*" & yarnConsumptionInfoDic("polyesterPercentage") & "%"
    totalConsumptionRange.Range("f" & rowTracker + 11).Style = "Comma"
    totalConsumptionRange.Range("f" & rowTracker + 11 & ":j" & rowTracker + 11).Merge
    totalConsumptionRange.Range("f" & rowTracker + 11 & ":j" & rowTracker + 11).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("s" & rowTracker + 11).value = yarnConsumptionInfoDic("pfd")
    totalConsumptionRange.Range("s" & rowTracker + 11 & ":y" & rowTracker + 11).Merge
    totalConsumptionRange.Range("s" & rowTracker + 11 & ":y" & rowTracker + 11).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 11).value = "PFD"
    totalConsumptionRange.Range("ag" & rowTracker + 11 & ":am" & rowTracker + 11).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 11 & ":am" & rowTracker + 11).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("a" & rowTracker + 12).value = "Spandex"
    totalConsumptionRange.Range("a" & rowTracker + 12 & ":d" & rowTracker + 12).Merge
    totalConsumptionRange.Range("a" & rowTracker + 12 & ":d" & rowTracker + 12).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("e" & rowTracker + 12).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("f" & rowTracker + 12).FormulaR1C1 = "=R[-4]C[-3]*" & yarnConsumptionInfoDic("spandexPercentage") & "%"
    totalConsumptionRange.Range("f" & rowTracker + 12).Style = "Comma"
    totalConsumptionRange.Range("f" & rowTracker + 12 & ":j" & rowTracker + 12).Merge
    totalConsumptionRange.Range("f" & rowTracker + 12 & ":j" & rowTracker + 12).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("s" & rowTracker + 12).value = yarnConsumptionInfoDic("ecru")
    totalConsumptionRange.Range("s" & rowTracker + 12 & ":y" & rowTracker + 12).Merge
    totalConsumptionRange.Range("s" & rowTracker + 12 & ":y" & rowTracker + 12).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 12).value = "ECRU"
    totalConsumptionRange.Range("ag" & rowTracker + 12 & ":am" & rowTracker + 12).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 12 & ":am" & rowTracker + 12).BorderAround, Weight:=xlThin


End Function

Private Function addPiInfoSourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus As Object) As Object
        'add PI data to UP issuing status

    Application.ScreenUpdating = False
        
    Dim piReportWb As Workbook
    Dim piReportWs As Worksheet
    Set piReportWb = Workbooks.Open(ActiveWorkbook.path & Application.PathSeparator & "PIReport.xlsx")
    Set piReportWs = piReportWb.Worksheets(1)

    piReportWs.AutoFilterMode = False
        
    Dim temp As Variant
    temp = piReportWs.Range("A4").CurrentRegion.value

    piReportWb.Close SaveChanges:=False
        
    Dim commercialFileNoDic As Object
    Set commercialFileNoDic = CreateObject("Scripting.Dictionary")
    
    Dim dicKey As Variant

    For Each dicKey In sourceDataAsDicUpIssuingStatus.keys

        If Not commercialFileNoDic.Exists(sourceDataAsDicUpIssuingStatus(dicKey)("CommercialFileNo")) Then

                'take unique commercial file name as dictionary key & assign a new dictionary
            commercialFileNoDic.Add sourceDataAsDicUpIssuingStatus(dicKey)("CommercialFileNo"), CreateObject("Scripting.Dictionary")

        End If

    Next dicKey

    Dim tempFabricCodeDicAsCommercialFile As Object
    
    Dim propertiesArr, propertiesValArr As Variant
    
    ReDim propertiesArr(1 To UBound(temp, 2))
    ReDim propertiesValArr(1 To UBound(temp, 2))
    
    Dim i, j As Long
    
    For j = 1 To UBound(temp, 2)
            'take first row as properties
        If IsEmpty(temp(1, j)) Then
            propertiesArr(j) = "column" & j 'empty key conflict handle
        Else  
            propertiesArr(j) = temp(1, j)
        End If

    Next j
    
    For i = 1 To UBound(temp)
        
        If commercialFileNoDic.Exists(temp(i, 3)) Then
        
            For j = 1 To UBound(temp, 2)
                propertiesValArr(j) = temp(i, j)
            Next j
        
            Set tempFabricCodeDicAsCommercialFile = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)
                    
            commercialFileNoDic(temp(i, 3)).Add commercialFileNoDic(temp(i, 3)).Count + 1, tempFabricCodeDicAsCommercialFile
        
        End If

    Next i

    For Each dicKey In sourceDataAsDicUpIssuingStatus.keys

        sourceDataAsDicUpIssuingStatus(dicKey).Add "fabricsInfo", commercialFileNoDic(sourceDataAsDicUpIssuingStatus(dicKey)("CommercialFileNo"))

    Next dicKey
            
    Set addPiInfoSourceDataAsDicUpIssuingStatus = sourceDataAsDicUpIssuingStatus

End Function

Private Function addYarnConsumptionInfoSourceDataAsDicUpIssuingStatus(sourceDataAsDicUpIssuingStatus As Object) As Object
        'add yarn consumption data to UP issuing status

        'below list of all needed key and value example
        ' yarnConsumptionInfoDic("weight") = 10.75 'should be dynamic
        ' yarnConsumptionInfoDic("width") = 66.5 'should be dynamic
        ' yarnConsumptionInfoDic("fabricQty") = 5000 'should be dynamic
        ' yarnConsumptionInfoDic("black") = "Black" 'add as PI color
        ' yarnConsumptionInfoDic("mercerizationBlack") = "Mercerization(Black)" 'add as PI
        ' yarnConsumptionInfoDic("indigo") = "Indigo" 'add as PI color
        ' yarnConsumptionInfoDic("mercerizationIndigo") = "Mercerization(Indigo)" 'add as PI
        ' yarnConsumptionInfoDic("toppingBottoming") = "Topping/ Bottoming" 'add as PI color
        ' yarnConsumptionInfoDic("mercerizationtoppingBottoming") = "Mercerization(Topping/ Bottoming)" 'add as PI
        ' yarnConsumptionInfoDic("overDying") = "Over Dying" 'add as PI
        ' yarnConsumptionInfoDic("mercerizationoverDying") = "Mercerization(Over Dying)" 'add as PI
        ' yarnConsumptionInfoDic("cottonPercentage") = 90 'should be dynamic
        ' yarnConsumptionInfoDic("coating") = "Coating" 'add as PI
        ' yarnConsumptionInfoDic("polyesterPercentage") = 5 'should be dynamic
        ' yarnConsumptionInfoDic("pfd") = "PFD" 'add as PI
        ' yarnConsumptionInfoDic("spandexPercentage") = 5 'should be dynamic
        ' yarnConsumptionInfoDic("ecru") = "ECRU" 'add as PI

    Dim dicKey As Variant
    Dim innerDicKey As Variant

    Dim fabricQtyInYds As Variant

    Dim yarnPercentage As Object

    Dim isBlack As Boolean
    Dim isIndigo As Boolean
    Dim isToppingBottoming As Boolean
    Dim isMercerization As Boolean
    Dim isOverDying As Boolean
    Dim isCoating As Boolean
    Dim isPfd As Boolean
    Dim isEcru As Boolean

    Dim sumFractionOfMtrQty As Variant

    For Each dicKey In sourceDataAsDicUpIssuingStatus.keys

            'add yarn consumption dictionary
        sourceDataAsDicUpIssuingStatus(dicKey).Add "yarnConsumptionInfo", CreateObject("Scripting.Dictionary")

        sumFractionOfMtrQty = 0 'reset

        For Each innerDicKey In sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo").keys

                'reset all variable 
            isBlack = False
            isIndigo = False
            isToppingBottoming = False
            isMercerization = False
            isOverDying = False
            isCoating = False
            isPfd = False
            isEcru = False

                'add inner dictionary & use dictionary key as dictionary count
            sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Add sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count + 1, CreateObject("Scripting.Dictionary")

                'add weight
                '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
            sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("weight") = _
                sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Weight")

                'add width
                '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
            sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("width") = _
                Application.Run("yarnConsumption.fabricWidthCalculation", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Width"))

                'add fabricQty
                '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
            If sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Unit") = "MTR" Then
                sumFractionOfMtrQty = sumFractionOfMtrQty + (sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("PIQty") * 1.0936132983 - _
                    Round(sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("PIQty") * 1.0936132983))
                fabricQtyInYds = Round(sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("PIQty") * 1.0936132983)

                If sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo").Count = innerDicKey Then
                    fabricQtyInYds = fabricQtyInYds + Round(sumFractionOfMtrQty)
                End If
            Else
                fabricQtyInYds = sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("PIQty")
            End If
            sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("fabricQty") = _
                fabricQtyInYds


            isBlack = Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Color"), _
                "(black)|(vanta)", True, True, True)
            isIndigo = Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Color"), _
                "(indigo)|(blue)", True, True, True)
            isToppingBottoming = Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Color"), _
                "(topping)|(bottoming)|(bi.?color)", True, True, True)
            isMercerization = Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Finished"), _
                "mercerize", True, True, True)
            isOverDying = Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Finished"), _
                "over", True, True, True)
            isCoating = Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Finished"), _
                "coated", True, True, True)
            isPfd = Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Finished"), _
                "pfd|bleach", True, True, True)
            isEcru = Application.Run("general_utility_functions.isStrPatternExist", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Color"), _
                "ecru", True, True, True)

            If (isToppingBottoming) Or (isBlack And isIndigo) Then
                    'add toppingBottoming
                    '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("toppingBottoming") = _
                    "Topping/ Bottoming"

                If isMercerization Or isOverDying Then 'if over dying then mercerize automatic taken
                        'add mercerizationtoppingBottoming
                        '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                    sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("mercerizationtoppingBottoming") = _
                        "Mercerization(Topping/ Bottoming)"
                End If
                
            ElseIf isBlack Then
                    'add black
                    '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("black") = _
                    "Black"

                If isMercerization Or isOverDying Then 'if over dying then mercerize automatic taken
                        'add mercerizationBlack
                        '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                    sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("mercerizationBlack") = _
                        "Mercerization(Black)"

                End If

            ElseIf isIndigo Then
                    'add indigo
                    '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("indigo") = _
                    "Indigo"

                If isMercerization Or isOverDying Then 'if over dying then mercerize automatic taken
                        'add mercerizationIndigo
                        '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                    sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("mercerizationIndigo") = _
                        "Mercerization(Indigo)"

                End If

            End If

            If isOverDying Then
                    'add overDying
                    '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("overDying") = _
                    "Over Dying"

                'no use case mercerizationoverDying
            End If

            If isCoating Then
                    'add coating
                    '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("coating") = _
                    "Coating"
            End If

            If isPfd Or ((Not isToppingBottoming) And (Not isBlack) And (Not isIndigo)) Then
                    'add pfd
                    '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("pfd") = _
                    "PFD"
            End If

            If isEcru Or (Not isToppingBottoming And Not isBlack And Not isIndigo) Then
                    'add ecru
                    '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
                sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("ecru") = _
                    "ECRU"

            End If

            Set yarnPercentage = Application.Run("yarnConsumption.calculateYarnPercentage", sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Composition"))

                'add cottonPercentage
                '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
            sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("cottonPercentage") = _
                yarnPercentage("cotton")

                'add polyesterPercentage
                '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
            sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("polyesterPercentage") = _
                yarnPercentage("polyester")

                'add spandexPercentage
                '***inner dictionary key must be same as dictionary key of "yarnConsumptionInfoDic" of function parameter  "yarnConsumptionInformationPutToProvidedWs"
            sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(sourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count)("spandexPercentage") = _
                yarnPercentage("spandex")

            If (isBlack = False) And (isIndigo = False) And (isToppingBottoming = False) And (isOverDying = True) Then
                Dim answer As VbMsgBoxResult
                answer = MsgBox("Color (" & sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Color") & _
                    ") in PI No. " & sourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("PINo") & _
                    " (Over Dying) process exist, But now calculating without any color. Do you agree?", _
                    vbYesNo + vbCritical + vbDefaultButton2, "Finished process conflict")
                
                If answer = vbNo Then
                    
                    'over dying exist, but color assuming Nill, which not possible, current color to be add with actual group
                    Err.Raise vbObjectError + 1000, , "Customs Err to stop procedure"
                    
                End If
            End If

        Next innerDicKey

    Next dicKey
    
    Set addYarnConsumptionInfoSourceDataAsDicUpIssuingStatus = sourceDataAsDicUpIssuingStatus

End Function


Private Function dealWithConsumptionSheet(consumptionWorksheet As Worksheet, withYarnConsumptionInfosourceDataAsDicUpIssuingStatus As Object)

    Application.ScreenUpdating = False

    Dim dicKey As Variant
    Dim innerDicKey As Variant
    Dim rowTracker As Long
    Dim outerLoopCounter As Long

    Dim topRow, bottomRow As Long

    topRow = 4
    bottomRow = consumptionWorksheet.Cells.Find("TOTAL", LookAt:=xlPart).Row - 4

    Dim totalConsumptionRange As Range
    Set totalConsumptionRange = consumptionWorksheet.Range("A" & topRow & ":" & "AM" & bottomRow)

    totalConsumptionRange.Rows("2:" & totalConsumptionRange.Rows.Count - 1).EntireRow.Delete

    With totalConsumptionRange.Rows("1")
        .Clear
        .Interior.ColorIndex = 2
        .RowHeight = 15
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 9
    End With

    Dim rowsCountForInsert As Long

    rowsCountForInsert = 0

    For Each dicKey In withYarnConsumptionInfosourceDataAsDicUpIssuingStatus.keys

            'count total yarnConsumptionInfo dictionary
        rowsCountForInsert = rowsCountForInsert + withYarnConsumptionInfosourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").Count

    Next dicKey

        'for each consumption portion 13 & bellow 1, for each buyer 1, and extra 2 rows for bottom
    rowsCountForInsert = (rowsCountForInsert * 13) + rowsCountForInsert + withYarnConsumptionInfosourceDataAsDicUpIssuingStatus.Count + 2

    totalConsumptionRange.Rows("2:" & rowsCountForInsert).EntireRow.Insert

    outerLoopCounter = 0
    rowTracker = 1

    For Each dicKey In withYarnConsumptionInfosourceDataAsDicUpIssuingStatus.keys

        outerLoopCounter = outerLoopCounter + 1 'for buyer Sl. No.

        If withYarnConsumptionInfosourceDataAsDicUpIssuingStatus.Count = 1 Then

            totalConsumptionRange.Range("a" & rowTracker).value = withYarnConsumptionInfosourceDataAsDicUpIssuingStatus(dicKey)("NameofBuyers")

            With totalConsumptionRange.Range("a" & rowTracker & ":y" & rowTracker)
                .Merge
                .Interior.ColorIndex = 6
                .HorizontalAlignment = xlLeft
                .RowHeight = 20
            End With

        Else

            totalConsumptionRange.Range("a" & rowTracker).value = outerLoopCounter & ") " & withYarnConsumptionInfosourceDataAsDicUpIssuingStatus(dicKey)("NameofBuyers")

            With totalConsumptionRange.Range("a" & rowTracker & ":y" & rowTracker)
                .Merge
                .Interior.ColorIndex = 6
                .HorizontalAlignment = xlLeft
                .RowHeight = 20
            End With

        End If


        rowTracker = rowTracker + 1

        For Each innerDicKey In withYarnConsumptionInfosourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo").keys

            Application.Run "yarnConsumption.yarnConsumptionInformationPutToProvidedWs", totalConsumptionRange, rowTracker, _
                withYarnConsumptionInfosourceDataAsDicUpIssuingStatus(dicKey)("yarnConsumptionInfo")(innerDicKey)

            rowTracker = rowTracker + 14

        Next innerDicKey

    Next dicKey

    Application.ScreenUpdating = True

End Function

Private Function fabricWidthCalculation(width As Variant) As Variant

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.MultiLine = True

    Dim storeWidthForValidation As Variant
    storeWidthForValidation = width 'just for use bottom validation section
        
    Dim result As Variant
            
    regEx.Pattern = "\-"
    
    width = regEx.Replace(width, "/")
    
    regEx.Pattern = "\s\/"
    
    width = regEx.Replace(width, "/")
    
    regEx.Pattern = "\/\s"
    
    width = regEx.Replace(width, "/")
    
    regEx.Pattern = "(\d+\.\d+)|(\d+)"
    
    Dim extractedWidth As Variant
    Set extractedWidth = regEx.Execute(width)
    
    If extractedWidth.Count > 1 Then
    
        Dim allWidthArr() As Variant
        ReDim allWidthArr(0 To extractedWidth.Count - 1)
        
        Dim allWidthArrIterator As Long
        For allWidthArrIterator = 0 To extractedWidth.Count - 1
            allWidthArr(allWidthArrIterator) = CDbl(extractedWidth(allWidthArrIterator).Value)
        Next allWidthArrIterator
            
        Dim infoAboutWidth As Object
        Set infoAboutWidth = CreateObject("Scripting.Dictionary")
        
        infoAboutWidth("allWidthArrLengthGreaterThanOne") = UBound(allWidthArr) > 1
        regEx.Pattern = "((\d+\.\d+)|(\d+))\/((\d+\.\d+)|(\d+))"
        infoAboutWidth("slashBesideDigit") = regEx.test(width)
        
        If infoAboutWidth("slashBesideDigit") Then
        
            infoAboutWidth.Add "slashExtrac", regEx.Execute(width)
            
            infoAboutWidth("slashOneTime") = infoAboutWidth("slashExtrac").Count = 1
            infoAboutWidth("slashTwoTime") = infoAboutWidth("slashExtrac").Count = 2
            
        End If
        
        If infoAboutWidth("allWidthArrLengthGreaterThanOne") Then
            Dim maxTwo As Variant
            ' maxTwo = Application.Run("yarnConsumptionModule.FindMaxTwoNumbers", allWidthArr)
            Set maxTwo = Application.Run("Sorting_Algorithms.FindMaxTwoNumbers", allWidthArr)

        End If
        
        If infoAboutWidth("allWidthArrLengthGreaterThanOne") And infoAboutWidth("slashBesideDigit") Then
            
            Dim leftWidth, rightWidth As Variant
            
            If infoAboutWidth("slashOneTime") Then
        
                regEx.Pattern = "(\d+\.\d+\/)|(\d+\/)"
                        
                Set leftWidth = regEx.Execute(infoAboutWidth("slashExtrac").Item(0))
                leftWidth = Replace(leftWidth.Item(0), "/", "")
        
                regEx.Pattern = "(\/\d+\.\d+)|(\/\d+)"
        
                Set rightWidth = regEx.Execute(infoAboutWidth("slashExtrac").Item(0))
                rightWidth = Replace(rightWidth.Item(0), "/", "")
        
            ElseIf infoAboutWidth("slashTwoTime") Then
                regEx.Pattern = "(\d+\.\d+\/)|(\d+\/)"
            
                Set leftWidth = regEx.Execute(infoAboutWidth("slashExtrac").Item(1))
                leftWidth = Replace(leftWidth.Item(0), "/", "")
        
                regEx.Pattern = "(\/\d+\.\d+)|(\/\d+)"
        
                Set rightWidth = regEx.Execute(infoAboutWidth("slashExtrac").Item(1))
                rightWidth = Replace(rightWidth.Item(0), "/", "")
            End If
    
        End If
        
        If infoAboutWidth("slashOneTime") And extractedWidth.Count = 2 Then
        
            result = (CDbl(extractedWidth(0).Value) + CDbl(extractedWidth(1).Value)) / 2
        
        ElseIf infoAboutWidth("slashOneTime") And extractedWidth.Count = 3 Then

        
        Dim slashOneTimeExclude As Variant
        slashOneTimeExclude = Application.Run("general_utility_functions.ExcludeElements", allWidthArr, Array(leftWidth, rightWidth))
        
        If slashOneTimeExclude(0) = maxTwo("firstMax") Then
            result = slashOneTimeExclude(0)
            
        ElseIf CDbl(leftWidth) = maxTwo("secondMax") And CDbl(rightWidth) = maxTwo("firstMax") Then
            result = (CDbl(leftWidth) + CDbl(rightWidth)) / 2
        End If
        
        ElseIf infoAboutWidth("slashTwoTime") Then
            Dim slashTwoTimeBesideNumberStr As Variant
            slashTwoTimeBesideNumberStr = infoAboutWidth("slashExtrac")(0).Value & " " & infoAboutWidth("slashExtrac")(1).Value
            
            regEx.Pattern = "(\d+\.\d+)|(\d+)"
            
            Dim slashTwoTimeBesideNumberExtract As Variant
            Set slashTwoTimeBesideNumberExtract = regEx.Execute(slashTwoTimeBesideNumberStr)
            
            Dim slashTwoTimeBesideNumberExtractArr() As Variant
            ReDim slashTwoTimeBesideNumberExtractArr(0 To slashTwoTimeBesideNumberExtract.Count - 1)
            
            Dim slashTwoTimeBesideNumberExtractArrIterator As Long
            For slashTwoTimeBesideNumberExtractArrIterator = 0 To slashTwoTimeBesideNumberExtract.Count - 1
                slashTwoTimeBesideNumberExtractArr(slashTwoTimeBesideNumberExtractArrIterator) = CDbl(slashTwoTimeBesideNumberExtract(slashTwoTimeBesideNumberExtractArrIterator).Value)
            Next slashTwoTimeBesideNumberExtractArrIterator
                 
            Dim slashTwoTimeBesideNumberExtractArrMaxTwo As Variant
            ' slashTwoTimeBesideNumberExtractArrMaxTwo = Application.Run("yarnConsumptionModule.FindMaxTwoNumbers", slashTwoTimeBesideNumberExtractArr)
            Set slashTwoTimeBesideNumberExtractArrMaxTwo = Application.Run("Sorting_Algorithms.FindMaxTwoNumbers", slashTwoTimeBesideNumberExtractArr)

            
            If CDbl(leftWidth) = slashTwoTimeBesideNumberExtractArrMaxTwo("secondMax") And CDbl(rightWidth) = slashTwoTimeBesideNumberExtractArrMaxTwo("firstMax") Then
            result = (CDbl(leftWidth) + CDbl(rightWidth)) / 2
            End If
            
        End If
    
    Else
    
        result = CDbl(extractedWidth(0).Value)
        
    End If

        'final width validation
    storeWidthForValidation = Replace(storeWidthForValidation, " ", "") 'replace space
    
    Dim tempDict As Object
    Set tempDict = CreateObject("Scripting.Dictionary")
    
    Dim sortedArrForValidation As Variant
   
    Dim extractWidthForValidation As Object
     
    regEx.pattern = "(\d+\.\d+)|(\d+)"
    
    Set extractWidthForValidation = regEx.Execute(storeWidthForValidation)

    Dim Match As Object

    For Each Match In extractWidthForValidation
        tempDict(Match.value) = Match.value
    Next
    
    sortedArrForValidation = Application.Run("Sorting_Algorithms.BubbleSort", tempDict.keys)  'sort ascending order
    
    If Not Application.Run("utilityFunction.isCompareValuesLessThanProvidedValue", result, sortedArrForValidation(UBound(sortedArrForValidation)), 0.51) Then

        MsgBox "Width " & storeWidthForValidation & " Calculation Error"
        Exit Function
        
    End If
    
    fabricWidthCalculation = result
    
  End Function

Private Function calculateYarnPercentage(fabricComposition As Variant) As Object

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.pattern = "(\d+\.\d+\%)|(\d+\%)"
    regEx.MultiLine = True

    Dim fabricCompositionStoreForMsg As Variant
    fabricCompositionStoreForMsg = fabricComposition

    Dim yarnGroup As Object
    Set yarnGroup = CreateObject("Scripting.Dictionary")
    
    Set yarnGroup = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", yarnGroup, "cotton", _
        Array("cotton", _
        "pcw", _
        "PreCW", _
        "PIW", _
        "RCS", _
        "ocs", _
        "bci", _
        "LINEN", _
        "grs", _
        "Hemp Yarn"))
    
    Set yarnGroup = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", yarnGroup, "polyester", _
        Array("polyester", _
        "Tencel", _
        "poly", _
        "Viscose", _
        "Rayon"))
        
    Set yarnGroup = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", yarnGroup, "spandex", _
       Array("spandex", _
       "elastane", _
       "Lycra", _
       "ELASTOMULTISTER", _
       "elastomultiester"))
    
    Dim sumPercentageAsYarnGroup As Object
    Set sumPercentageAsYarnGroup = CreateObject("Scripting.Dictionary")
    
    sumPercentageAsYarnGroup("cotton") = 0
    sumPercentageAsYarnGroup("polyester") = 0
    sumPercentageAsYarnGroup("spandex") = 0

    fabricComposition = Replace(fabricComposition, " ", "") 'replace space
    
    Dim regExReturnedObjectExtractPercentage As Variant
    Set regExReturnedObjectExtractPercentage = regEx.Execute(fabricComposition)

    Dim percentageReplaceWithComma As Variant
    percentageReplaceWithComma = Replace(fabricComposition, ",", "")  'replace previous comma
    percentageReplaceWithComma = regEx.Replace(percentageReplaceWithComma, ",")  'replace percentage portion with commma
    percentageReplaceWithComma = Replace(percentageReplaceWithComma, ",", "", 1, 1) ' replace first comma only

    Dim extractYarnCategory As Variant
    extractYarnCategory = Split(percentageReplaceWithComma, ",")

    Dim extractYarnCategoryArrayIterator As Long
    Dim dictKey As Variant

    For extractYarnCategoryArrayIterator = 0 To UBound(extractYarnCategory)

        For Each dictKey In yarnGroup.keys
            
            If Application.Run("general_utility_functions.isStrPatternExist", Application.Run("general_utility_functions.RemoveInvalidChars", extractYarnCategory(extractYarnCategoryArrayIterator)), dictKey, True, True, True) Then
            
                sumPercentageAsYarnGroup(yarnGroup(dictKey)) = sumPercentageAsYarnGroup(yarnGroup(dictKey)) _
                + Round(CDec(Replace(regExReturnedObjectExtractPercentage.Item(extractYarnCategoryArrayIterator), "%", "")), 2) 'some time type conflict, handle error type converted
                    
                    Exit For

            End If
    
        Next

    Next extractYarnCategoryArrayIterator
    
    If sumPercentageAsYarnGroup("cotton") + sumPercentageAsYarnGroup("polyester") + sumPercentageAsYarnGroup("spandex") <> 100 Then
        MsgBox fabricCompositionStoreForMsg & Chr(10) & "Above Total sum of Yarn percentage not 100 may be new yarn group exist"
        Debug.Print fabricCompositionStoreForMsg 'print Immediate window for copy
        Exit Function
    End If
    
    Set calculateYarnPercentage = sumPercentageAsYarnGroup

End Function


Private Function validateCommercialFileQtyAndUnit(withPiInfosourceDataAsDicUpIssuingStatus As Object)

    Dim validateFileqty As Object
    Set validateFileqty = CreateObject("Scripting.Dictionary")

    Dim validateFileUnit As Object
    Set validateFileUnit = CreateObject("Scripting.Dictionary")

    Dim tempSum As Variant
    tempSum = 0

    Dim qtyUnitFromPiInfo As String
    Dim qtyUnitFromUpIssuingStatus As String

    Dim dicKey As Variant
    Dim innerDicKey As Variant

    For Each dicKey In withPiInfosourceDataAsDicUpIssuingStatus.keys

        For Each innerDicKey In withPiInfosourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo").keys

            tempSum = tempSum + withPiInfosourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("PIQty")

            qtyUnitFromPiInfo = withPiInfosourceDataAsDicUpIssuingStatus(dicKey)("fabricsInfo")(innerDicKey)("Unit")
            
        Next innerDicKey

        If withPiInfosourceDataAsDicUpIssuingStatus(dicKey)("QuantityofFabricsYdsMtr") <> tempSum Then
                'add commercial file
            validateFileqty(withPiInfosourceDataAsDicUpIssuingStatus(dicKey)("CommercialFileNo")) = withPiInfosourceDataAsDicUpIssuingStatus(dicKey)("CommercialFileNo")

        End If

        tempSum = 0 'reset

            'Qty. unit pick form up issuing status
        If Right(withPiInfosourceDataAsDicUpIssuingStatus(dicKey)("qtyNumberFormat"), 5) = """Mtr""" Then

            qtyUnitFromUpIssuingStatus = "MTR"
        Else

            qtyUnitFromUpIssuingStatus = "YDS"

        End If

        If qtyUnitFromUpIssuingStatus <> qtyUnitFromPiInfo Then
                'add commercial file
            validateFileUnit(withPiInfosourceDataAsDicUpIssuingStatus(dicKey)("CommercialFileNo")) = withPiInfosourceDataAsDicUpIssuingStatus(dicKey)("CommercialFileNo")

        End If

        qtyUnitFromPiInfo = "" 'reset
        qtyUnitFromUpIssuingStatus = "" 'reset

    Next dicKey

    Dim tempQtyMsg As String
    tempQtyMsg = "May be bellow commercial file include multiple LC Amnd, Make unique file manully by add extension with file No. in both sheet up issuing status & PI info" & Chr(10)

    If validateFileqty.Count > 0 Then

        For Each dicKey In validateFileqty.keys

            tempQtyMsg = tempQtyMsg & validateFileqty(dicKey) & Chr(10)

        Next dicKey
        
        MsgBox tempQtyMsg
        Err.Raise vbObjectError + 1000, , "Customs Err to stop procedure"

    End If

    Dim tempQtyUnitMsg As String
    tempQtyUnitMsg = "Bellow commercial file Qty. unit mismatch in up issuing status & PI info, Please corrected" & Chr(10)

    If validateFileUnit.Count > 0 Then

        For Each dicKey In validateFileUnit.keys

            tempQtyUnitMsg = tempQtyUnitMsg & validateFileUnit(dicKey) & Chr(10)

        Next dicKey
        
        MsgBox tempQtyUnitMsg
        Err.Raise vbObjectError + 1000, , "Customs Err to stop procedure"

    End If

End Function