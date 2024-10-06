Attribute VB_Name = "readUp"
Option Explicit

Private Function readUpAsDict(upWs As Worksheet) As Object

    Dim upAsDict As Object
    Set upAsDict = CreateObject("Scripting.Dictionary")
        
    Dim isAfterCustomsAct2023Formate As Boolean
    isAfterCustomsAct2023Formate = False ' Initialize the flag
    
    If Application.Run("utilityFunction.DoesStringExistInWorksheets", "8|  Avg`vwb Gjwmi weeiY t", upWs) Then

        Dim topRow As Long
        topRow = upWs.Cells.Find("8|  Avg`vwb Gjwmi weeiY t", LookAt:=xlPart).Row + 1

        If Left(upWs.Cells(topRow, 3).value, 4) = "Gjwm" Then
            
            isAfterCustomsAct2023Formate = True

        End If

    End If

    upAsDict.Add "upClause1", Application.Run("readUp.upClause1AsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause6", Application.Run("readUp.upClause6AsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause7", Application.Run("readUp.upClause7AsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause8", Application.Run("readUp.upClause8AsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause9", Application.Run("readUp.upClause9AsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause11", Application.Run("readUp.upClause11AsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause12a", Application.Run("readUp.upClause12aAsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause12bFabrics", Application.Run("readUp.upClause12bFabricsAsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause12bGarments", Application.Run("readUp.upClause12bGarmentsAsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause13", Application.Run("readUp.upClause13AsDict", upWs, isAfterCustomsAct2023Formate)
    upAsDict.Add "upClause14", Application.Run("readUp.upClause14AsDict", upWs, isAfterCustomsAct2023Formate)
    
    
    Set readUpAsDict = upAsDict
    
End Function

Private Function upClause1AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause1AsDict As Object
    Set clause1AsDict = CreateObject("Scripting.Dictionary")

    Dim curentUpNo As Variant
    curentUpNo = Application.Run("helperFunctionGetData.upNoFromProvidedWs", upWs)
    
    clause1AsDict("upNo") = curentUpNo

    Set upClause1AsDict = clause1AsDict
    
End Function

Private Function upClause6AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause6AsDict As Object
    Set clause6AsDict = CreateObject("Scripting.Dictionary")

    Dim clause6Range As Object

    If isAfterCustomsAct2023Formate Then

        Set clause6Range = Application.Run("helperFunctionGetRangeObject.upClause6BuyerinformationRangeObjectFromProvidedWs", upWs)

    Else

        Set clause6Range = Application.Run("previousFormatRelatedFun.upClause6BuyerinformationRangeObjectFromProvidedWsPrevFormat", upWs)

    End If

    Dim clause6Arr As Variant
    Dim regExObj As Object
    Dim i As Long

    clause6Arr = clause6Range.Value
    Set regExObj = Application.Run("general_utility_functions.createRegExObj", "^\d\s*\)", True, True, True)

    For i = LBound(clause6Arr) To UBound(clause6Arr)

        clause6AsDict.Add clause6AsDict.Count + 1, Trim(regExObj.Replace(Trim(clause6Arr(i, 14)), ""))

    Next i
        
    Set upClause6AsDict = clause6AsDict
    
End Function

Private Function upClause7AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause7AsDict As Object
    Set clause7AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause7AsDict = clause7AsDict
    
End Function

Private Function upClause8AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause8AsDict As Object
    Set clause8AsDict = CreateObject("Scripting.Dictionary")

    If isAfterCustomsAct2023Formate Then

        Set clause8AsDict = Application.Run("general_utility_functions.upClause8InformationFromProvidedWs", upWs)

    Else

        Set clause8AsDict = Application.Run("previousFormatRelatedFun.upClause8InformationFromProvidedWsPrevFormat", upWs)

    End If
        
    Set upClause8AsDict = clause8AsDict
    
End Function

Private Function upClause9AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause9AsDict As Object
    Set clause9AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause9AsDict = clause9AsDict
    
End Function

Private Function upClause11AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause11AsDict As Object
    Set clause11AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause11AsDict = clause11AsDict
    
End Function

Private Function upClause12aAsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause12aAsDict As Object
    Set clause12aAsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause12aAsDict = clause12aAsDict
    
End Function

Private Function upClause12bFabricsAsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause12bFabricsAsDict As Object
    Set clause12bFabricsAsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause12bFabricsAsDict = clause12bFabricsAsDict
    
End Function

Private Function upClause12bGarmentsAsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause12bGarmentsAsDict As Object
    Set clause12bGarmentsAsDict = CreateObject("Scripting.Dictionary")
    
    Set upClause12bGarmentsAsDict = clause12bGarmentsAsDict
    
End Function

Private Function upClause13AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause13AsDict As Object
    Set clause13AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause13AsDict = clause13AsDict
    
End Function

Private Function upClause14AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause14AsDict As Object
    Set clause14AsDict = CreateObject("Scripting.Dictionary")
    
    Set upClause14AsDict = clause14AsDict
    
End Function

