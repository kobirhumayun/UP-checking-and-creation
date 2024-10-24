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

    Dim clause7Range As Object

    If isAfterCustomsAct2023Formate Then

        Set clause7Range = Application.Run("helperFunctionGetRangeObject.upClause7LcinformationRangeObjectFromProvidedWs", upWs)

    Else

        Set clause7Range = Application.Run("previousFormatRelatedFun.upClause7LcinformationRangeObjectFromProvidedWsPrevFormat", upWs)

    End If

    Dim isGarments As Boolean
    Dim clause7Arr As Variant
    Dim dicKey As Variant
    Dim lcFieldVal As String
    Dim bankFieldVal As String
    Dim qtyTopFieldVal As Variant 'type Variant to check empty value
    Dim qtyBottomFieldVal As Variant 'type Variant to check empty value
    Dim lcValueTopFieldVal As Variant 'type Variant to check empty value
    Dim lcValueBottomFieldVal As Variant 'type Variant to check empty value
    Dim tempRegExReturnedObj As Object
    Dim i As Long

    clause7Arr = clause7Range.Value

    isGarments = Application.Run("general_utility_functions.isStrPatternExist", clause7Arr(2, 17), "garments", True, True, True)

    For i = (LBound(clause7Arr) + 1) To (UBound(clause7Arr) - 1) Step 2 'exclude first & last rows

        clause7AsDict.Add clause7AsDict.Count + 1, CreateObject("Scripting.Dictionary")

        clause7AsDict(clause7AsDict.Count).Add "isGarments", isGarments

        If isAfterCustomsAct2023Formate Then

            lcFieldVal = clause7Arr(i, 3)
            bankFieldVal = clause7Arr(i, 10)

        Else

            lcFieldVal = clause7Arr(i, 4)
            bankFieldVal = clause7Arr(i, 12)

        End If

        Dim lcFieldDict As Object
        Set lcFieldDict = Application.Run("readUp.ExtractLCField", lcFieldVal)

        For Each dicKey In lcFieldDict.keys
            
            clause7AsDict(clause7AsDict.Count).Add dicKey, lcFieldDict(dicKey)

        Next dicKey

        clause7AsDict(clause7AsDict.Count).Add "bankName", bankFieldVal

        clause7AsDict(clause7AsDict.Count).Add "shipmentDate", CDate(clause7Arr(i, 16))
        clause7AsDict(clause7AsDict.Count).Add "expiryDate", CDate(clause7Arr(i + 1, 16))

        qtyTopFieldVal = clause7Arr(i, 18)
        qtyBottomFieldVal = clause7Arr(i + 1, 18)

        If isGarments Then

            clause7AsDict(clause7AsDict.Count).Add "garmentsQty", qtyTopFieldVal

            If Application.Run("general_utility_functions.isStrPatternExist", qtyBottomFieldVal, "mtr", True, True, True) Then

                clause7AsDict(clause7AsDict.Count).Add "isFabQtyExistInMtr", True

                Set tempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", qtyBottomFieldVal, ".+mtr$", True, True, True)
                clause7AsDict(clause7AsDict.Count).Add "fabricsQtyInMtr", CDec(Replace(tempRegExReturnedObj(0), "Mtr", ""))

                Set tempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", qtyBottomFieldVal, ".+yds$", True, True, True)
                clause7AsDict(clause7AsDict.Count).Add "fabricsQtyInYds", CDec(Replace(tempRegExReturnedObj(0), "Yds", ""))

            Else
                    
                clause7AsDict(clause7AsDict.Count).Add "isFabQtyExistInMtr", False
                clause7AsDict(clause7AsDict.Count).Add "fabricsQtyInYds", CDec(qtyBottomFieldVal)

            End If

        Else

            If Application.Run("general_utility_functions.isStrPatternExist", qtyTopFieldVal, "mtr", True, True, True) Then

                clause7AsDict(clause7AsDict.Count).Add "isFabQtyExistInMtr", True

                clause7AsDict(clause7AsDict.Count).Add "fabricsQtyInMtr", CDec(Replace(qtyTopFieldVal, "Mtr", ""))

                clause7AsDict(clause7AsDict.Count).Add "fabricsQtyInYds", CDec(qtyBottomFieldVal)

            Else

                If Not IsEmpty(qtyTopFieldVal) And IsEmpty(qtyBottomFieldVal) Then
                    
                    clause7AsDict(clause7AsDict.Count).Add "isFabQtyExistInMtr", False
                    clause7AsDict(clause7AsDict.Count).Add "fabricsQtyInYds", CDec(qtyTopFieldVal)

                Else

                    MsgBox upWs.Name & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "Bottom Qty. field not Empty but Qty in Yds"
                    Exit Function

                End If

            End If

        End If

        lcValueTopFieldVal = clause7Arr(i, 20)
        lcValueBottomFieldVal = clause7Arr(i + 1, 20)

        If Application.Run("general_utility_functions.isStrPatternExist", lcValueTopFieldVal, "euro", True, True, True) Then

            clause7AsDict(clause7AsDict.Count).Add "isLcValueExistInEuro", True

            clause7AsDict(clause7AsDict.Count).Add "lcValueInEuro", CDec(Replace(lcValueTopFieldVal, "Euro", ""))

            clause7AsDict(clause7AsDict.Count).Add "lcValueInUsd", CDec(lcValueBottomFieldVal)

        Else

            If Not IsEmpty(lcValueTopFieldVal) And IsEmpty(lcValueBottomFieldVal) Then
                
                clause7AsDict(clause7AsDict.Count).Add "isLcValueExistInEuro", False
                clause7AsDict(clause7AsDict.Count).Add "lcValueInUsd", CDec(lcValueTopFieldVal)

            Else

                MsgBox upWs.Name & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "Bottom lc Value field not Empty but value in USD"
                Exit Function

            End If

        End If

        Dim mlcExpIpLeftFieldVal As Variant
        Dim mlcExpIpRightFieldVal As Variant
        Dim tempDict As Object

        mlcExpIpLeftFieldVal = clause7Arr(i, 22)
        mlcExpIpRightFieldVal = clause7Arr(i, 25)

        If isAfterCustomsAct2023Formate Then

            If Application.Run("general_utility_functions.isStrPatternExist", mlcExpIpLeftFieldVal, "ip", True, True, True) Then
                ' EPZ
                clause7AsDict(clause7AsDict.Count).Add "isExistIp", True
                clause7AsDict(clause7AsDict.Count).Add "isExistExp", True
                clause7AsDict(clause7AsDict.Count).Add "isExistMlc", False

                Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", mlcExpIpLeftFieldVal, "ip.+\n?\d{2}\/\d{2}\/\d{4}", "ip")

                If tempDict.Count > 0 Then
                    clause7AsDict(clause7AsDict.Count).Add "ip", tempDict
                Else
                    MsgBox "#1000" & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "IP not found in UP clause 7"
                End If

                Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", mlcExpIpLeftFieldVal, "exp.+\n?\d{2}\/\d{2}\/\d{4}", "exp")

                If tempDict.Count > 0 Then
                    clause7AsDict(clause7AsDict.Count).Add "exp", tempDict
                Else
                    MsgBox "#1001" & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "EXP not found in UP clause 7"
                End If

            ElseIf Application.Run("general_utility_functions.isStrPatternExist", mlcExpIpLeftFieldVal, "exp", True, True, True) Then
                ' direct
                clause7AsDict(clause7AsDict.Count).Add "isExistIp", False
                clause7AsDict(clause7AsDict.Count).Add "isExistExp", True
                clause7AsDict(clause7AsDict.Count).Add "isExistMlc", False

                Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", mlcExpIpLeftFieldVal, "exp.+\n?\d{2}\/\d{2}\/\d{4}", "exp")

                If tempDict.Count > 0 Then
                    clause7AsDict(clause7AsDict.Count).Add "exp", tempDict
                Else
                    MsgBox "#1002" & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "EXP not found in UP clause 7"
                End If

            Else
                ' Deem
                clause7AsDict(clause7AsDict.Count).Add "isExistIp", False
                clause7AsDict(clause7AsDict.Count).Add "isExistExp", False
                clause7AsDict(clause7AsDict.Count).Add "isExistMlc", True

                Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", mlcExpIpLeftFieldVal, ".+\n?\d{2}\/\d{2}\/\d{4}", "mlc")

                If tempDict.Count > 0 Then
                    clause7AsDict(clause7AsDict.Count).Add "mlc", tempDict
                Else
                    MsgBox "#1003" & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "MLC not found in UP clause 7"
                End If

            End If

        Else
            ' previous UP format
            Dim isIpAndExp As Boolean
            Dim isOnlyExp As Boolean

            isIpAndExp = False
            isOnlyExp = False

            isIpAndExp = Application.Run("general_utility_functions.isStrPatternExist", mlcExpIpLeftFieldVal, "\d+\/\d{6}\/\d{4}.*\n?\d{2}\/\d{2}\/\d{4}", True, True, True) _
                And Application.Run("general_utility_functions.isStrPatternExist", mlcExpIpRightFieldVal, ".+\n?\d{2}\/\d{2}\/\d{4}", True, True, True)
            
            isOnlyExp = Application.Run("general_utility_functions.isStrPatternExist", mlcExpIpLeftFieldVal, "\d+\/\d{6}\/\d{4}.*\n?\d{2}\/\d{2}\/\d{4}", True, True, True) _
                And IsEmpty(mlcExpIpRightFieldVal)

            If isIpAndExp Then
                ' EPZ
                clause7AsDict(clause7AsDict.Count).Add "isExistIp", True
                clause7AsDict(clause7AsDict.Count).Add "isExistExp", True
                clause7AsDict(clause7AsDict.Count).Add "isExistMlc", False

                Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", mlcExpIpRightFieldVal, ".+\n?\d{2}\/\d{2}\/\d{4}", "ip")

                If tempDict.Count > 0 Then
                    clause7AsDict(clause7AsDict.Count).Add "ip", tempDict
                Else
                    MsgBox "#1004" & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "IP not found in UP clause 7"
                End If

                Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", mlcExpIpLeftFieldVal, "\d+\/\d{6}\/\d{4}.*\n?\d{2}\/\d{2}\/\d{4}", "exp")

                If tempDict.Count > 0 Then
                    clause7AsDict(clause7AsDict.Count).Add "exp", tempDict
                Else
                    MsgBox "#1005" & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "EXP not found in UP clause 7"
                End If

            ElseIf isOnlyExp Then
                ' direct
                clause7AsDict(clause7AsDict.Count).Add "isExistIp", False
                clause7AsDict(clause7AsDict.Count).Add "isExistExp", True
                clause7AsDict(clause7AsDict.Count).Add "isExistMlc", False

                Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", mlcExpIpLeftFieldVal, "\d+\/\d{6}\/\d{4}.*\n?\d{2}\/\d{2}\/\d{4}", "exp")

                If tempDict.Count > 0 Then
                    clause7AsDict(clause7AsDict.Count).Add "exp", tempDict
                Else
                    MsgBox "#1006" & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "EXP not found in UP clause 7"
                End If

            Else
                ' Deem
                clause7AsDict(clause7AsDict.Count).Add "isExistIp", False
                clause7AsDict(clause7AsDict.Count).Add "isExistExp", False
                clause7AsDict(clause7AsDict.Count).Add "isExistMlc", True

                Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", mlcExpIpLeftFieldVal, ".+\n?\d{2}\/\d{2}\/\d{4}", "mlc")

                If tempDict.Count > 0 Then
                    clause7AsDict(clause7AsDict.Count).Add "mlc", tempDict
                Else
                    MsgBox "#1007" & Chr(10) & clause7AsDict(clause7AsDict.Count)("lcNo") & Chr(10) & "MLC not found in UP clause 7"
                End If

            End If

        End If

    Next i

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

' ========utility function=========

Private Function ExtractLCField(lcFieldVal As String) As Object

    Dim lcFieldDict As Object
    Set lcFieldDict = CreateObject("Scripting.Dictionary")
    Dim tempRegExReturnedObj As Object



    lcFieldDict.Add "lcNo", Application.Run("general_utility_functions.ExtractFirstLineWithRegex", lcFieldVal)

    Set tempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", lcFieldVal, "\d{2}\/\d{2}\/\d{4}", True, True, True)
    lcFieldDict.Add "lcDt", tempRegExReturnedObj(0) 'first occurrence

    If Application.Run("general_utility_functions.isStrPatternExist", lcFieldVal, "amnd", True, True, True) Then

        lcFieldDict.Add "isLcAmndExist", True

        Set tempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", lcFieldVal, "amnd\-\d+", True, True, True)
        Set tempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", tempRegExReturnedObj(0), "\d+$", True, True, True)
        lcFieldDict.Add "lcAmndNo", CInt(tempRegExReturnedObj(0)) 'exclude left zero

        Set tempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", lcFieldVal, "\d{2}\/\d{2}\/\d{4}", True, True, True)
        lcFieldDict.Add "lcAmndDt", tempRegExReturnedObj(1) 'second occurrence

    Else

        lcFieldDict.Add "isLcAmndExist", False

    End If

    If Application.Run("general_utility_functions.isStrPatternExist", lcFieldVal, "\(.+\)", True, True, True) Then

        lcFieldDict.Add "isDcNoExist", True
        Set tempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", lcFieldVal, "\(.+\)", True, True, True)
        Set tempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", tempRegExReturnedObj(0), "\d+", True, True, True)
        lcFieldDict.Add "dcNo", tempRegExReturnedObj(0)

    Else

        lcFieldDict.Add "isDcNoExist", False

    End If
    
    Set ExtractLCField = lcFieldDict
    
End Function

Private Function MlcUdIpExpAndDtExtractor(receivedStr As String, regExPattern As String, mlcUdIpExpKeyName As String) As Object

    Dim mlcUdIpExpDict As Object
    Set mlcUdIpExpDict = CreateObject("Scripting.Dictionary")
    Dim tempRegExReturnedObj As Object
    Dim innerTempRegExReturnedObj As Object
    Dim match As Object
    Dim tempStr As String
    Dim tempDateStr As String
    Dim tempMlcUdIpExpStr As String
    Dim removedAllInvalidChrForKeys As String

    Set tempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", receivedStr, regExPattern, True, True, True)

    For Each match In tempRegExReturnedObj

        tempStr =  Trim(Replace(match.Value, Chr(10), ""))
        Set innerTempRegExReturnedObj = Application.Run("general_utility_functions.regExReturnedObj", tempStr, "\d{2}\/\d{2}\/\d{4}$", True, True, True)

        tempDateStr = innerTempRegExReturnedObj(0)
        tempMlcUdIpExpStr = Trim(Replace(tempStr, tempDateStr, ""))

        removedAllInvalidChrForKeys = Application.Run("general_utility_functions.RemoveInvalidChars", tempMlcUdIpExpStr)

        mlcUdIpExpDict.Add removedAllInvalidChrForKeys, CreateObject("Scripting.Dictionary")
        mlcUdIpExpDict(removedAllInvalidChrForKeys).Add mlcUdIpExpKeyName, tempMlcUdIpExpStr
        mlcUdIpExpDict(removedAllInvalidChrForKeys).Add "date", tempDateStr

    Next
    
    Set MlcUdIpExpAndDtExtractor = mlcUdIpExpDict
    
End Function

