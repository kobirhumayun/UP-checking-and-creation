Attribute VB_Name = "yarnConsumption"
Option Explicit


Private Function yarnConsumptionInformationPutToProvidedWs(totalConsumptionRange As Range, rowTracker As Long, yarnConsumptionInfoDic As Object)
    'this function put yarnConsumption information to provided worksheet

    totalConsumptionRange.Range("a" & rowTracker).value = "Weight :"
    totalConsumptionRange.Range("a" & rowTracker & ":c" & rowTracker).Merge

        'weight
    totalConsumptionRange.Range("d" & rowTracker).value = 10.5 'to be dynamic
    totalConsumptionRange.Range("d" & rowTracker & ":e" & rowTracker).Merge

    totalConsumptionRange.Range("f" & rowTracker).value = "OZ/YD2"
    totalConsumptionRange.Range("f" & rowTracker & ":g" & rowTracker).Merge

    totalConsumptionRange.Range("i" & rowTracker).value = "Width :"
    totalConsumptionRange.Range("i" & rowTracker & ":k" & rowTracker).Merge

        'Width
    totalConsumptionRange.Range("l" & rowTracker).value = 66.5 'to be dynamic
    totalConsumptionRange.Range("l" & rowTracker & ":n" & rowTracker).Merge

    totalConsumptionRange.Range("o" & rowTracker).value = "Inch"
    totalConsumptionRange.Range("o" & rowTracker & ":p" & rowTracker).Merge

    totalConsumptionRange.Range("r" & rowTracker).value = "Qty :"
    totalConsumptionRange.Range("r" & rowTracker & ":s" & rowTracker).Merge

        'Qty.
    totalConsumptionRange.Range("t" & rowTracker).value = 5000 'to be dynamic
    totalConsumptionRange.Range("t" & rowTracker & ":v" & rowTracker).Merge

    totalConsumptionRange.Range("w" & rowTracker).value = "Yds"
    totalConsumptionRange.Range("w" & rowTracker & ":x" & rowTracker).Merge


    totalConsumptionRange.Range("b" & rowTracker + 2).value = "="

        'put formula to take weight
    totalConsumptionRange.Range("c" & rowTracker + 2).FormulaR1C1 = "=R[-2]C[1]"
    totalConsumptionRange.Range("c" & rowTracker + 2 & ":d" & rowTracker + 2).Merge

    totalConsumptionRange.Range("e" & rowTracker + 2).value = "x"

        'put formula to take width
    totalConsumptionRange.Range("f" & rowTracker + 2).FormulaR1C1 = "=R[-2]C[6]"

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

    totalConsumptionRange.Range("s" & rowTracker + 2).value = "Black" 'to be dynamic
    totalConsumptionRange.Range("s" & rowTracker + 2 & ":y" & rowTracker + 2).Merge
    totalConsumptionRange.Range("s" & rowTracker + 2 & ":y" & rowTracker + 2).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 2).value = "Black"
    totalConsumptionRange.Range("ag" & rowTracker + 2 & ":am" & rowTracker + 2).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 2 & ":am" & rowTracker + 2).BorderAround, Weight:=xlThin


    totalConsumptionRange.Range("s" & rowTracker + 3).value = "Mercerization(Black)" 'to be dynamic
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

    totalConsumptionRange.Range("s" & rowTracker + 4).value = "Indigo" 'to be dynamic
    totalConsumptionRange.Range("s" & rowTracker + 4 & ":y" & rowTracker + 4).Merge
    totalConsumptionRange.Range("s" & rowTracker + 4 & ":y" & rowTracker + 4).BorderAround, Weight:=xlThin

    totalConsumptionRange.Range("ag" & rowTracker + 4).value = "Indigo"
    totalConsumptionRange.Range("ag" & rowTracker + 4 & ":am" & rowTracker + 4).Merge
    totalConsumptionRange.Range("ag" & rowTracker + 4 & ":am" & rowTracker + 4).BorderAround, Weight:=xlThin






























End Function
