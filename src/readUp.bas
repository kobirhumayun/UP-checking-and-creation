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

        clause7AsDict(clause7AsDict.Count).Add "shipmentDate", CStr(clause7Arr(i, 16))
        clause7AsDict(clause7AsDict.Count).Add "expiryDate", CStr(clause7Arr(i + 1, 16))

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

            If Application.Run("general_utility_functions.isStrPatternExist", mlcExpIpLeftFieldVal, "ip\:", True, True, True) Then
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

            ElseIf Application.Run("general_utility_functions.isStrPatternExist", mlcExpIpLeftFieldVal, "exp\:", True, True, True) Then
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

    Dim upClause9StockinformationRangeObject As Object
    Dim upClause9Val As Variant

    Set upClause9StockinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause9StockinformationRangeObjectFromProvidedWs", upWs)

    upClause9Val = upClause9StockinformationRangeObject.Value

    clause9AsDict.Add "yarnImport", CreateObject("Scripting.Dictionary")
    clause9AsDict.Add "yarnLocal", CreateObject("Scripting.Dictionary")
    clause9AsDict.Add "dyes", CreateObject("Scripting.Dictionary")
    clause9AsDict.Add "chemicalsImport", CreateObject("Scripting.Dictionary")
    clause9AsDict.Add "chemicalsLocal", CreateObject("Scripting.Dictionary")
    clause9AsDict.Add "stretchWrappingFilm", CreateObject("Scripting.Dictionary")

    clause9AsDict("yarnImport").Add "previousDue", upClause9Val(1 ,14)
    clause9AsDict("yarnLocal").Add "previousDue", upClause9Val(2 ,14)
    clause9AsDict("dyes").Add "previousDue", upClause9Val(3 ,14)
    clause9AsDict("chemicalsImport").Add "previousDue", upClause9Val(4 ,14)
    clause9AsDict("chemicalsLocal").Add "previousDue", upClause9Val(5 ,14)
    clause9AsDict("stretchWrappingFilm").Add "previousDue", upClause9Val(6 ,14)

    clause9AsDict("yarnImport").Add "currentImport", upClause9Val(1 ,16)
    clause9AsDict("yarnLocal").Add "currentImport", upClause9Val(2 ,16)
    clause9AsDict("dyes").Add "currentImport", upClause9Val(3 ,16)
    clause9AsDict("chemicalsImport").Add "currentImport", upClause9Val(4 ,16)
    clause9AsDict("chemicalsLocal").Add "currentImport", upClause9Val(5 ,16)
    clause9AsDict("stretchWrappingFilm").Add "currentImport", upClause9Val(6 ,16)
        
    clause9AsDict("yarnImport").Add "sumPreviousDueCurrentImport", upClause9Val(1 ,18)
    clause9AsDict("yarnLocal").Add "sumPreviousDueCurrentImport", upClause9Val(2 ,18)
    clause9AsDict("dyes").Add "sumPreviousDueCurrentImport", upClause9Val(3 ,18)
    clause9AsDict("chemicalsImport").Add "sumPreviousDueCurrentImport", upClause9Val(4 ,18)
    clause9AsDict("chemicalsLocal").Add "sumPreviousDueCurrentImport", upClause9Val(5 ,18)
    clause9AsDict("stretchWrappingFilm").Add "sumPreviousDueCurrentImport", upClause9Val(6 ,18)

    clause9AsDict("yarnImport").Add "previousUsed", upClause9Val(1 ,20)
    clause9AsDict("yarnLocal").Add "previousUsed", upClause9Val(2 ,20)
    clause9AsDict("dyes").Add "previousUsed", upClause9Val(3 ,20)
    clause9AsDict("chemicalsImport").Add "previousUsed", upClause9Val(4 ,20)
    clause9AsDict("chemicalsLocal").Add "previousUsed", upClause9Val(5 ,20)
    clause9AsDict("stretchWrappingFilm").Add "previousUsed", upClause9Val(6 ,20)

    clause9AsDict("yarnImport").Add "currentStock", upClause9Val(1 ,22)
    clause9AsDict("yarnLocal").Add "currentStock", upClause9Val(2 ,22)
    clause9AsDict("dyes").Add "currentStock", upClause9Val(3 ,22)
    clause9AsDict("chemicalsImport").Add "currentStock", upClause9Val(4 ,22)
    clause9AsDict("chemicalsLocal").Add "currentStock", upClause9Val(5 ,22)
    clause9AsDict("stretchWrappingFilm").Add "currentStock", upClause9Val(6 ,22)

    clause9AsDict("yarnImport").Add "usedInThisUp", upClause9Val(1 ,24)
    clause9AsDict("yarnLocal").Add "usedInThisUp", upClause9Val(2 ,24)
    clause9AsDict("dyes").Add "usedInThisUp", upClause9Val(3 ,24)
    clause9AsDict("chemicalsImport").Add "usedInThisUp", upClause9Val(4 ,24)
    clause9AsDict("chemicalsLocal").Add "usedInThisUp", upClause9Val(5 ,24)
    clause9AsDict("stretchWrappingFilm").Add "usedInThisUp", upClause9Val(6 ,24)

    clause9AsDict("yarnImport").Add "remainingQty", upClause9Val(1 ,26)
    clause9AsDict("yarnLocal").Add "remainingQty", upClause9Val(2 ,26)
    clause9AsDict("dyes").Add "remainingQty", upClause9Val(3 ,26)
    clause9AsDict("chemicalsImport").Add "remainingQty", upClause9Val(4 ,26)
    clause9AsDict("chemicalsLocal").Add "remainingQty", upClause9Val(5 ,26)
    clause9AsDict("stretchWrappingFilm").Add "remainingQty", upClause9Val(6 ,26)

    Set upClause9AsDict = clause9AsDict
    
End Function

Private Function upClause11AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause11AsDict As Object
    Set clause11AsDict = CreateObject("Scripting.Dictionary")
    Dim clause11Arr As Variant
    Dim buyerName As String

    Dim upClause11UdExpIpinformationRangeObject As Range
    Set upClause11UdExpIpinformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause11UdExpIpinformationRangeObjectFromProvidedWs", upWs)

    clause11Arr = upClause11UdExpIpinformationRangeObject.Value

    Dim i As Long

    For i = LBound(clause11Arr) To UBound(clause11Arr) - 1

        clause11AsDict.Add clause11AsDict.Count + 1, CreateObject("Scripting.Dictionary")

        If isAfterCustomsAct2023Formate Then

            buyerName = clause11Arr(i, 3)

        Else

            buyerName = clause11Arr(i, 4)

        End If

        clause11AsDict(clause11AsDict.Count).Add "buyerName", buyerName

        Dim udExpIp As Variant
        Dim tempDict As Object

        udExpIp = clause11Arr(i, 17)

        If Application.Run("general_utility_functions.isStrPatternExist", udExpIp, "ip", True, True, True) Then
            ' EPZ
            clause11AsDict(clause11AsDict.Count).Add "isExistIp", True
            clause11AsDict(clause11AsDict.Count).Add "isExistExp", True
            clause11AsDict(clause11AsDict.Count).Add "isExistUd", False

            Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", udExpIp, "ip.+\n?\d{2}\/\d{2}\/\d{4}", "ip")

            If tempDict.Count > 0 Then
                clause11AsDict(clause11AsDict.Count).Add "ip", tempDict
            Else
                MsgBox "#1008" & Chr(10) & "Sl. " & clause11AsDict.Count & Chr(10) & "IP not found in UP clause 11"
            End If

            Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", udExpIp, "exp.+\n?\d{2}\/\d{2}\/\d{4}", "exp")

            If tempDict.Count > 0 Then
                clause11AsDict(clause11AsDict.Count).Add "exp", tempDict
            Else
                MsgBox "#1009" & Chr(10) & "Sl. " & clause11AsDict.Count & Chr(10) & "EXP not found in UP clause 11"
            End If

        ElseIf Application.Run("general_utility_functions.isStrPatternExist", udExpIp, "exp", True, True, True) Then
            ' direct
            clause11AsDict(clause11AsDict.Count).Add "isExistIp", False
            clause11AsDict(clause11AsDict.Count).Add "isExistExp", True
            clause11AsDict(clause11AsDict.Count).Add "isExistUd", False

            Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", udExpIp, "exp.+\n?\d{2}\/\d{2}\/\d{4}", "exp")

            If tempDict.Count > 0 Then
                clause11AsDict(clause11AsDict.Count).Add "exp", tempDict
            Else
                MsgBox "#1010" & Chr(10) & "Sl. " & clause11AsDict.Count & Chr(10) & "EXP not found in UP clause 11"
            End If

        Else
            ' Deem
            clause11AsDict(clause11AsDict.Count).Add "isExistIp", False
            clause11AsDict(clause11AsDict.Count).Add "isExistExp", False
            clause11AsDict(clause11AsDict.Count).Add "isExistUd", True

            Set tempDict = Application.Run("readUp.MlcUdIpExpAndDtExtractor", udExpIp, ".+\n?\d{2}\/\d{2}\/\d{4}", "ud")

            If tempDict.Count > 0 Then
                clause11AsDict(clause11AsDict.Count).Add "ud", tempDict
            Else
                MsgBox "#1011" & Chr(10) & "Sl. " & clause11AsDict.Count & Chr(10) & "MLC not found in UP clause 11"
            End If

        End If

        clause11AsDict(clause11AsDict.Count).Add "fabricWidth", clause11Arr(i, 23)
        clause11AsDict(clause11AsDict.Count).Add "fabricWeight", clause11Arr(i, 25)
        clause11AsDict(clause11AsDict.Count).Add "fabricQty", clause11Arr(i, 26)

    Next i

    Set upClause11AsDict = clause11AsDict
    
End Function

Private Function upClause12aAsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause12aAsDict As Object
    Set clause12aAsDict = CreateObject("Scripting.Dictionary")
    Dim clause12aArr As Variant
    Dim buyerName As String

    Dim upClause12AYarnConsumptionInformationRangeObject As Range

    If isAfterCustomsAct2023Formate Then

        Set upClause12AYarnConsumptionInformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause12AYarnConsumptioninformationRangeObjectFromProvidedWs", upWs)

    Else

        Set upClause12AYarnConsumptionInformationRangeObject = Application.Run("previousFormatRelatedFun.upClause12AYarnConsumptioninformationRangeObjectFromProvidedWsPrevFormat", upWs)

    End If

    clause12aArr = upClause12AYarnConsumptionInformationRangeObject.Value

    Dim i As Long

    For i = LBound(clause12aArr) To UBound(clause12aArr) - 1

        If Not IsEmpty(clause12aArr(i, 3)) Then

            clause12aAsDict.Add clause12aAsDict.Count + 1, CreateObject("Scripting.Dictionary")
            buyerName = clause12aArr(i, 3)

        End If

        clause12aAsDict(clause12aAsDict.Count).Add clause12aAsDict(clause12aAsDict.Count).Count + 1, CreateObject("Scripting.Dictionary")

        clause12aAsDict(clause12aAsDict.Count)(clause12aAsDict(clause12aAsDict.Count).Count).Add "buyerName", buyerName
        clause12aAsDict(clause12aAsDict.Count)(clause12aAsDict(clause12aAsDict.Count).Count).Add "garmentsQty", clause12aArr(i, 18)
        clause12aAsDict(clause12aAsDict.Count)(clause12aAsDict(clause12aAsDict.Count).Count).Add "fabricQty", clause12aArr(i, 19)
        clause12aAsDict(clause12aAsDict.Count)(clause12aAsDict(clause12aAsDict.Count).Count).Add "yarnConPerKg", clause12aArr(i, 21)
        clause12aAsDict(clause12aAsDict.Count)(clause12aAsDict(clause12aAsDict.Count).Count).Add "totalYarnUsed", clause12aArr(i, 23)
        clause12aAsDict(clause12aAsDict.Count)(clause12aAsDict(clause12aAsDict.Count).Count).Add "overconsumptionPercentage", clause12aArr(i, 25)
        clause12aAsDict(clause12aAsDict.Count)(clause12aAsDict(clause12aAsDict.Count).Count).Add "totalYarnUsedWithOverconsumption", clause12aArr(i, 26)

    Next i
        
    Set upClause12aAsDict = clause12aAsDict
    
End Function

Private Function upClause12bFabricsAsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause12bFabricsAsDict As Object
    Set clause12bFabricsAsDict = CreateObject("Scripting.Dictionary")
    Dim clause12bFabArr As Variant
        
    Dim upClause12BYarnConsumptionInformationRangeObject As Range
    Set upClause12BYarnConsumptionInformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause12BChemicalDyesConsumptioninformationRangeObjectFromProvidedWs", upWs)

    clause12bFabArr = upClause12BYarnConsumptionInformationRangeObject.Value

    clause12bFabricsAsDict.Add "grandTotalYarn", clause12bFabArr(1, 7)

    clause12bFabricsAsDict.Add "buyerName", CreateObject("Scripting.Dictionary")
    clause12bFabricsAsDict.Add "quantityOfGoodsUsedInProduction", CreateObject("Scripting.Dictionary")
    clause12bFabricsAsDict.Add "rawMaterials", CreateObject("Scripting.Dictionary")

    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "ropeDenimFabricsDyedBlack", clause12bFabArr(12, 7)
    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "ropeDenimFabricsDyedBlackMercerization", clause12bFabArr(18, 7)
    
    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "ropeDenimFabricsDyedIndigo", clause12bFabArr(32, 7)
    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "ropeDenimFabricsDyedIndigoMercerization", clause12bFabArr(38, 7)

    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "ropeDenimFabricsDyed", clause12bFabArr(53, 7)
    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "ropeDenimFabricsDyedMercerization", clause12bFabArr(61, 7)

    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "denimFabricsOverDyedSolidDyed", clause12bFabArr(72, 7)
    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "denimFabricsOverDyedSolidDyedMercerization", clause12bFabArr(75, 7)

    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "denimFabricsCoatedAndPigment", clause12bFabArr(82, 7)
    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "denimFabricsPFDFinished", clause12bFabArr(92, 7)
    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "denimFabricsEcruFinished", clause12bFabArr(102, 7)
    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "denimFabricDyed", clause12bFabArr(107, 7)
    clause12bFabricsAsDict("quantityOfGoodsUsedInProduction").Add "denimFabricPacking", clause12bFabArr(111, 7)

    Dim i As Long
    Dim removedAllInvalidChrFromKeys As String

    For i = LBound(clause12bFabArr) + 1 To UBound(clause12bFabArr)

        If Not IsEmpty(clause12bFabArr(i, 2)) Then

            clause12bFabricsAsDict("buyerName").Add clause12bFabricsAsDict("buyerName").Count + 1, clause12bFabArr(i, 2)

        End If

        removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", clause12bFabArr(i, 16))   'remove all invalid characters for use dic keys

        clause12bFabricsAsDict("rawMaterials").Add removedAllInvalidChrFromKeys & "_Sl_" & i - 1, clause12bFabArr(i, 25)

    Next i

    Set upClause12bFabricsAsDict = clause12bFabricsAsDict
    
End Function

Private Function upClause12bGarmentsAsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause12bGarmentsAsDict As Object
    Set clause12bGarmentsAsDict = CreateObject("Scripting.Dictionary")
    Dim isGarments As Boolean
    Dim clause12bGarmentsArr As Variant

    Dim upClause12BGarmentsRangeObject As Range
    Set upClause12BGarmentsRangeObject = Application.Run("helperFunctionGetRangeObject.upClause12BGarmentsRangeObjectFromProvidedWs", upWs)

    clause12bGarmentsArr = upClause12BGarmentsRangeObject.Value

    Dim i As Long

    isGarments = False

    For i = LBound(clause12bGarmentsArr) To UBound(clause12bGarmentsArr)

        If Not IsEmpty(clause12bGarmentsArr(i, 11)) Then

            isGarments = True

        End If

    Next i

    clause12bGarmentsAsDict.Add "isGarments", isGarments

    clause12bGarmentsAsDict.Add "rawMaterials", CreateObject("Scripting.Dictionary")

    Dim removedAllInvalidChrFromKeys As String
    Dim typeOfWash As String

    If isAfterCustomsAct2023Formate Then
        
        For i = LBound(clause12bGarmentsArr) To UBound(clause12bGarmentsArr)

            If Not IsEmpty(clause12bGarmentsArr(i, 14)) Then

                removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", clause12bGarmentsArr(i, 14))   'remove all invalid characters for use dic keys
                typeOfWash = removedAllInvalidChrFromKeys
                clause12bGarmentsAsDict.Add typeOfWash, CreateObject("Scripting.Dictionary")

            End If

            If Not IsEmpty(clause12bGarmentsArr(i, 11)) Then

                removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", clause12bGarmentsArr(i, 2))   'remove all invalid characters for use dic keys

                clause12bGarmentsAsDict(typeOfWash).Add removedAllInvalidChrFromKeys, clause12bGarmentsArr(i, 11)

            End If

            removedAllInvalidChrFromKeys = Application.Run("general_utility_functions.RemoveInvalidChars", clause12bGarmentsArr(i, 15))   'remove all invalid characters for use dic keys

            clause12bGarmentsAsDict("rawMaterials").Add removedAllInvalidChrFromKeys & "_Sl_" & i, clause12bGarmentsArr(i, 25)

        Next i

    End If

    Set upClause12bGarmentsAsDict = clause12bGarmentsAsDict
    
End Function

Private Function upClause13AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause13AsDict As Object
    Set clause13AsDict = CreateObject("Scripting.Dictionary")

    Dim upClause13Val As Variant
    Dim upClause13InformationRangeObject As Range
    Set upClause13InformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause13UseRawMaterialsinformationRangeObjectFromProvidedWs", upWs)

    upClause13Val = upClause13InformationRangeObject.Value

    clause13AsDict.Add "yarnImport", CreateObject("Scripting.Dictionary")
    clause13AsDict.Add "yarnLocal", CreateObject("Scripting.Dictionary")
    clause13AsDict.Add "dyes", CreateObject("Scripting.Dictionary")
    clause13AsDict.Add "chemicalsImport", CreateObject("Scripting.Dictionary")
    clause13AsDict.Add "chemicalsLocal", CreateObject("Scripting.Dictionary")
    clause13AsDict.Add "stretchWrappingFilm", CreateObject("Scripting.Dictionary")

    clause13AsDict("yarnImport").Add "qty", upClause13Val(1 ,15)
    clause13AsDict("yarnLocal").Add "qty", upClause13Val(2 ,15)
    clause13AsDict("dyes").Add "qty", upClause13Val(3 ,15)
    clause13AsDict("chemicalsImport").Add "qty", upClause13Val(4 ,15)
    clause13AsDict("chemicalsLocal").Add "qty", upClause13Val(5 ,15)
    clause13AsDict("stretchWrappingFilm").Add "qty", upClause13Val(6 ,15)

    clause13AsDict("yarnImport").Add "value", upClause13Val(1 ,18)
    clause13AsDict("yarnLocal").Add "value", upClause13Val(2 ,18)
    clause13AsDict("dyes").Add "value", upClause13Val(3 ,18)
    clause13AsDict("chemicalsImport").Add "value", upClause13Val(4 ,18)
    clause13AsDict("chemicalsLocal").Add "value", upClause13Val(5 ,18)
    clause13AsDict("stretchWrappingFilm").Add "value", upClause13Val(6 ,18)

    Set upClause13AsDict = clause13AsDict
    
End Function

Private Function upClause14AsDict(upWs As Worksheet, isAfterCustomsAct2023Formate As Boolean) As Object

    Dim clause14AsDict As Object
    Set clause14AsDict = CreateObject("Scripting.Dictionary")

    Dim isGarments As Boolean
    Dim upClause14Val As Variant
    Dim upClause14InformationRangeObject As Range

    If isAfterCustomsAct2023Formate Then

        Set upClause14InformationRangeObject = Application.Run("helperFunctionGetRangeObject.upClause14RangeObjectFromProvidedWs", upWs)

    Else

        Set upClause14InformationRangeObject = Application.Run("previousFormatRelatedFun.upClause15RangeObjectFromProvidedWsPrevFormat", upWs)
                
    End If

    upClause14Val = upClause14InformationRangeObject.Value

    isGarments = False

    If Not IsEmpty(upClause14Val(1, 17)) Then

        isGarments = True

    End If

    clause14AsDict.Add "isGarments", isGarments

    If isGarments Then
        
        clause14AsDict.Add "garmentsQty", upClause14Val(1, 17)

    End If

    clause14AsDict.Add "fabricsQty", upClause14Val(1, 15)
    clause14AsDict.Add "exportValue", upClause14Val(2, 15)
    clause14AsDict.Add "rawMaterialsValue", upClause14Val(3, 15)
    clause14AsDict.Add "valueAddition", upClause14Val(4, 15)

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

Private Function saveUpDataAsJsonFromSelectedUpFile(jsonPath As String, initialUpPath As String)

    Application.ScreenUpdating = False

    Dim allUpDic As Object
    Set allUpDic = CreateObject("Scripting.Dictionary")
    Dim curentUpDict As Object

    Dim currentUpWb As Workbook
    Dim currentUpWs As Worksheet

    Dim answer As VbMsgBoxResult
    ' Display the message box with Yes and No buttons
    answer = MsgBox("Do you want to use previous calculated JSON text file", vbYesNo + vbQuestion, "JSON text file")
    Dim jsonPathArr As Variant

    ' Check which button the user clicked
    If answer = vbYes Then
        ' Code to execute if user clicks Yes

        jsonPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", jsonPath)  ' JSON file path

        If Not UBound(jsonPathArr) = 1 Then
            MsgBox "Please select only one JSON file"
            Exit Function
        End If

        Set allUpDic = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", jsonPathArr(1))

    End If

    Dim upPathArr As Variant
    upPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", initialUpPath)

    Dim i As Long
    For i = LBound(upPathArr) To UBound(upPathArr) ' create dictionary as UP

        Application.DisplayAlerts = False
        Set currentUpWb = Workbooks.Open(upPathArr(i))
        Set currentUpWs = currentUpWb.Worksheets(2)

        Set curentUpDict = Application.Run("readUp.readUpAsDict", currentUpWs)

        currentUpWb.Close SaveChanges:=False
        Application.DisplayAlerts = True

        If Not allUpDic.Exists(curentUpDict("upClause1")("upNo")) Then
            
            allUpDic.Add curentUpDict("upClause1")("upNo"), curentUpDict
        
        End If

    Next i

    Dim sortedAllCalculatedUp As Variant
    sortedAllCalculatedUp = Application.Run("Sorting_Algorithms.upSort", allUpDic.keys)

    Application.Run "JsonUtilityFunction.SaveDictionaryToJsonTextFile", allUpDic, jsonPath & Application.PathSeparator & _
        "UP-" & Replace(sortedAllCalculatedUp(LBound(sortedAllCalculatedUp)), "/", "-") & "-to-" & _
        Replace(sortedAllCalculatedUp(UBound(sortedAllCalculatedUp)), "/", "-") & "-all-clause-data" & ".json"

    Application.ScreenUpdating = True

    Dim uPSequenceStr As String
    uPSequenceStr = Application.Run("utilityFunction.upSequenceStrGenerator", allUpDic.keys, " -to- ", 10)

    MsgBox "Saved UP data as JSON" & Chr(10) & uPSequenceStr

End Function

Private Function loadUpDataFromJsonAndFormatedAsWriteToSheetAsUp(jsonPath As String, upNo As String) As Object

    Application.ScreenUpdating = False

    Dim allUpDic As Object
    Set allUpDic = CreateObject("Scripting.Dictionary")
    Dim curentUpDict As Object
    Dim curentUpAsWriteFormatDict As Object
    Set curentUpAsWriteFormatDict = CreateObject("Scripting.Dictionary")

    Dim jsonPathArr As Variant

    jsonPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", jsonPath)  ' JSON file path

    If Not UBound(jsonPathArr) = 1 Then
        MsgBox "Please select only one JSON file"
        Exit Function
    End If

    Set allUpDic = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", jsonPathArr(1))

    Set curentUpDict = allUpDic(upNo)

    curentUpAsWriteFormatDict.Add "upClause1", CreateObject("Scripting.Dictionary")
    curentUpAsWriteFormatDict("upClause1").Add curentUpAsWriteFormatDict("upClause1").Count + 1, curentUpDict("upClause1")("upNo")

    Dim outerKey, innerKey1, innerKey2 As Variant

    curentUpAsWriteFormatDict.Add "upClause6", CreateObject("Scripting.Dictionary")

    For Each outerKey In curentUpDict("upClause6").keys
    
        curentUpAsWriteFormatDict("upClause6").Add curentUpAsWriteFormatDict("upClause6").Count + 1, curentUpDict("upClause6")(outerKey)
        
    Next outerKey

    curentUpAsWriteFormatDict.Add "upClause7", CreateObject("Scripting.Dictionary")

    For Each outerKey In curentUpDict("upClause7").keys

        If Not IsEmpty(outerKey) Then 'to handle unknown empty dict key

            curentUpAsWriteFormatDict("upClause7").Add curentUpAsWriteFormatDict("upClause7").Count + 1, CreateObject("Scripting.Dictionary")

            For Each innerKey1 In curentUpDict("upClause7")(outerKey).keys

                If innerKey1 = "exp" Then

                    For Each innerKey2 In curentUpDict("upClause7")(outerKey)(innerKey1).keys

                        curentUpAsWriteFormatDict("upClause7")(curentUpAsWriteFormatDict("upClause7").Count).Add innerKey2 & "_exp", curentUpDict("upClause7")(outerKey)(innerKey1)(innerKey2)("exp")
                        curentUpAsWriteFormatDict("upClause7")(curentUpAsWriteFormatDict("upClause7").Count).Add innerKey2 & "_date", curentUpDict("upClause7")(outerKey)(innerKey1)(innerKey2)("date")

                    Next innerKey2

                ElseIf innerKey1 = "ip" Then

                    For Each innerKey2 In curentUpDict("upClause7")(outerKey)(innerKey1).keys

                        curentUpAsWriteFormatDict("upClause7")(curentUpAsWriteFormatDict("upClause7").Count).Add innerKey2 & "_ip", curentUpDict("upClause7")(outerKey)(innerKey1)(innerKey2)("ip")
                        curentUpAsWriteFormatDict("upClause7")(curentUpAsWriteFormatDict("upClause7").Count).Add innerKey2 & "_date", curentUpDict("upClause7")(outerKey)(innerKey1)(innerKey2)("date")

                    Next innerKey2

                ElseIf innerKey1 = "mlc" Then

                    For Each innerKey2 In curentUpDict("upClause7")(outerKey)(innerKey1).keys

                        curentUpAsWriteFormatDict("upClause7")(curentUpAsWriteFormatDict("upClause7").Count).Add innerKey2 & "_mlc", curentUpDict("upClause7")(outerKey)(innerKey1)(innerKey2)("mlc")
                        curentUpAsWriteFormatDict("upClause7")(curentUpAsWriteFormatDict("upClause7").Count).Add innerKey2 & "_date", curentUpDict("upClause7")(outerKey)(innerKey1)(innerKey2)("date")

                    Next innerKey2

                Else

                    curentUpAsWriteFormatDict("upClause7")(curentUpAsWriteFormatDict("upClause7").Count).Add innerKey1, curentUpDict("upClause7")(outerKey)(innerKey1)

                End If

            Next innerKey1

        End If

    Next outerKey

    curentUpAsWriteFormatDict.Add "upClause8", CreateObject("Scripting.Dictionary")

    For Each outerKey In curentUpDict("upClause8").keys

        curentUpAsWriteFormatDict("upClause8").Add curentUpAsWriteFormatDict("upClause8").Count + 1, CreateObject("Scripting.Dictionary")

        For Each innerKey1 In curentUpDict("upClause8")(outerKey).keys

            curentUpAsWriteFormatDict("upClause8")(curentUpAsWriteFormatDict("upClause8").Count).Add innerKey1, curentUpDict("upClause8")(outerKey)(innerKey1)

        Next innerKey1

    Next outerKey

    curentUpAsWriteFormatDict.Add "upClause9", CreateObject("Scripting.Dictionary")

    For Each outerKey In curentUpDict("upClause9").keys

        curentUpAsWriteFormatDict("upClause9").Add curentUpAsWriteFormatDict("upClause9").Count + 1, CreateObject("Scripting.Dictionary")

        For Each innerKey1 In curentUpDict("upClause9")(outerKey).keys

            curentUpAsWriteFormatDict("upClause9")(curentUpAsWriteFormatDict("upClause9").Count).Add outerKey & "_" & innerKey1, curentUpDict("upClause9")(outerKey)(innerKey1)

        Next innerKey1

    Next outerKey

    curentUpAsWriteFormatDict.Add "upClause11", CreateObject("Scripting.Dictionary")

    For Each outerKey In curentUpDict("upClause11").keys

        curentUpAsWriteFormatDict("upClause11").Add curentUpAsWriteFormatDict("upClause11").Count + 1, CreateObject("Scripting.Dictionary")

        For Each innerKey1 In curentUpDict("upClause11")(outerKey).keys

            If innerKey1 = "exp" Then

                For Each innerKey2 In curentUpDict("upClause11")(outerKey)(innerKey1).keys

                    curentUpAsWriteFormatDict("upClause11")(curentUpAsWriteFormatDict("upClause11").Count).Add innerKey2 & "_exp", curentUpDict("upClause11")(outerKey)(innerKey1)(innerKey2)("exp")
                    curentUpAsWriteFormatDict("upClause11")(curentUpAsWriteFormatDict("upClause11").Count).Add innerKey2 & "_date", curentUpDict("upClause11")(outerKey)(innerKey1)(innerKey2)("date")

                Next innerKey2

            ElseIf innerKey1 = "ip" Then

                For Each innerKey2 In curentUpDict("upClause11")(outerKey)(innerKey1).keys

                    curentUpAsWriteFormatDict("upClause11")(curentUpAsWriteFormatDict("upClause11").Count).Add innerKey2 & "_ip", curentUpDict("upClause11")(outerKey)(innerKey1)(innerKey2)("ip")
                    curentUpAsWriteFormatDict("upClause11")(curentUpAsWriteFormatDict("upClause11").Count).Add innerKey2 & "_date", curentUpDict("upClause11")(outerKey)(innerKey1)(innerKey2)("date")

                Next innerKey2

            ElseIf innerKey1 = "ud" Then

                For Each innerKey2 In curentUpDict("upClause11")(outerKey)(innerKey1).keys

                    curentUpAsWriteFormatDict("upClause11")(curentUpAsWriteFormatDict("upClause11").Count).Add innerKey2 & "_ud", curentUpDict("upClause11")(outerKey)(innerKey1)(innerKey2)("ud")
                    curentUpAsWriteFormatDict("upClause11")(curentUpAsWriteFormatDict("upClause11").Count).Add innerKey2 & "_date", curentUpDict("upClause11")(outerKey)(innerKey1)(innerKey2)("date")

                Next innerKey2

            Else

                curentUpAsWriteFormatDict("upClause11")(curentUpAsWriteFormatDict("upClause11").Count).Add innerKey1, curentUpDict("upClause11")(outerKey)(innerKey1)

            End If

        Next innerKey1

    Next outerKey

    curentUpAsWriteFormatDict.Add "upClause12a", CreateObject("Scripting.Dictionary")

    For Each outerKey In curentUpDict("upClause12a").keys

        For Each innerKey1 In curentUpDict("upClause12a")(outerKey).keys

            curentUpAsWriteFormatDict("upClause12a").Add curentUpAsWriteFormatDict("upClause12a").Count + 1, CreateObject("Scripting.Dictionary")

            For Each innerKey2 In curentUpDict("upClause12a")(outerKey)(innerKey1).keys

                curentUpAsWriteFormatDict("upClause12a")(curentUpAsWriteFormatDict("upClause12a").Count).Add innerKey2, curentUpDict("upClause12a")(outerKey)(innerKey1)(innerKey2)

            Next innerKey2

        Next innerKey1

    Next outerKey
    
    curentUpAsWriteFormatDict.Add "upClause12bFabrics", CreateObject("Scripting.Dictionary")

    curentUpAsWriteFormatDict("upClause12bFabrics").Add curentUpAsWriteFormatDict("upClause12bFabrics").Count + 1, CreateObject("Scripting.Dictionary")
    
    curentUpAsWriteFormatDict("upClause12bFabrics")(curentUpAsWriteFormatDict("upClause12bFabrics").Count).Add "grandTotalYarn", curentUpDict("upClause12bFabrics")("grandTotalYarn")

    For Each outerKey In curentUpDict("upClause12bFabrics")("buyerName").keys

        curentUpAsWriteFormatDict("upClause12bFabrics").Add curentUpAsWriteFormatDict("upClause12bFabrics").Count + 1, CreateObject("Scripting.Dictionary")
        curentUpAsWriteFormatDict("upClause12bFabrics")(curentUpAsWriteFormatDict("upClause12bFabrics").Count).Add "buyer_" & outerKey, curentUpDict("upClause12bFabrics")("buyerName")(outerKey)

    Next outerKey

    For Each outerKey In curentUpDict("upClause12bFabrics")("quantityOfGoodsUsedInProduction").keys

        curentUpAsWriteFormatDict("upClause12bFabrics").Add curentUpAsWriteFormatDict("upClause12bFabrics").Count + 1, CreateObject("Scripting.Dictionary")
        curentUpAsWriteFormatDict("upClause12bFabrics")(curentUpAsWriteFormatDict("upClause12bFabrics").Count).Add outerKey, curentUpDict("upClause12bFabrics")("quantityOfGoodsUsedInProduction")(outerKey)

    Next outerKey

    For Each outerKey In curentUpDict("upClause12bFabrics")("rawMaterials").keys

        curentUpAsWriteFormatDict("upClause12bFabrics").Add curentUpAsWriteFormatDict("upClause12bFabrics").Count + 1, CreateObject("Scripting.Dictionary")
        curentUpAsWriteFormatDict("upClause12bFabrics")(curentUpAsWriteFormatDict("upClause12bFabrics").Count).Add outerKey, curentUpDict("upClause12bFabrics")("rawMaterials")(outerKey)

    Next outerKey

    curentUpAsWriteFormatDict.Add "upClause12bGarments", CreateObject("Scripting.Dictionary")

    curentUpAsWriteFormatDict("upClause12bGarments").Add "isGarments", curentUpDict("upClause12bGarments")("isGarments")

    For Each outerKey In curentUpDict("upClause12bGarments").keys

        If TypeName(curentUpDict("upClause12bGarments")(outerKey)) = "Dictionary" Then
            
            For Each innerKey1 In curentUpDict("upClause12bGarments")(outerKey).keys

                curentUpAsWriteFormatDict("upClause12bGarments").Add outerKey & "_" & innerKey1, curentUpDict("upClause12bGarments")(outerKey)(innerKey1)

            Next innerKey1

        End If

    Next outerKey

    curentUpAsWriteFormatDict.Add "upClause13", CreateObject("Scripting.Dictionary")

    For Each outerKey In curentUpDict("upClause13").keys

        If TypeName(curentUpDict("upClause13")(outerKey)) = "Dictionary" Then
            
            For Each innerKey1 In curentUpDict("upClause13")(outerKey).keys

                curentUpAsWriteFormatDict("upClause13").Add outerKey & "_" & innerKey1, curentUpDict("upClause13")(outerKey)(innerKey1)

            Next innerKey1

        End If

    Next outerKey

    curentUpAsWriteFormatDict.Add "upClause14", CreateObject("Scripting.Dictionary")

    For Each outerKey In curentUpDict("upClause14").keys

        curentUpAsWriteFormatDict("upClause14").Add outerKey, curentUpDict("upClause14")(outerKey)

    Next outerKey

    Set loadUpDataFromJsonAndFormatedAsWriteToSheetAsUp = curentUpAsWriteFormatDict
    
End Function

Sub saveUpDataAsJson()

    Dim jsonPath As String
    Dim initialUpPath As String
    
    jsonPath = "D:\Temp\UP Draft\Draft 2024\json-all-up-clause"
    initialUpPath = "D:\Temp\UP Draft\Draft 2024"
    
    saveUpDataAsJsonFromSelectedUpFile jsonPath, initialUpPath

End Sub

Sub loadUpDataFromJsonAndWriteToSheetAsUptoVerify()

    Dim jsonPath As String
    jsonPath = "D:\Temp\UP Draft\Draft 2024\json-all-up-clause"

    Dim upNo As String
    upNo = InputBox("Please enter UP Number", "UP Number", "UP No.")

    Dim curentUpAsWriteFormatDict As Object
    Set curentUpAsWriteFormatDict = loadUpDataFromJsonAndFormatedAsWriteToSheetAsUp(jsonPath, upNo)

    Sheets.Add After:=Sheets(ActiveSheet.Name)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim outerKey As Variant
    Dim rowPtr As Long
    rowPtr = 1
    
    ws.Range("a" & rowPtr).value = "Clause 1"
    rowPtr = rowPtr + 1
    
    Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause1"), True, True, True
    rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause1").Count + 1

    ws.Range("a" & rowPtr).value = "Clause 6"
    rowPtr = rowPtr + 1
    
    Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause6"), True, True, True
    rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause6").Count + 1
        
    ws.Range("a" & rowPtr).value = "Clause 7"
    rowPtr = rowPtr + 1
    
    For Each outerKey In curentUpAsWriteFormatDict("upClause7").keys
    
        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause7")(outerKey), True, True, True
        rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause7")(outerKey).Count + 1
        
    Next outerKey
    
    ws.Range("a" & rowPtr).value = "Clause 8"
    rowPtr = rowPtr + 1
    
    For Each outerKey In curentUpAsWriteFormatDict("upClause8").keys
    
        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause8")(outerKey), True, True, True
        rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause8")(outerKey).Count + 1
        
    Next outerKey
    
    ws.Range("a" & rowPtr).value = "Clause 9"
    rowPtr = rowPtr + 1
    
    For Each outerKey In curentUpAsWriteFormatDict("upClause9").keys
    
        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause9")(outerKey), True, True, True
        rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause9")(outerKey).Count + 1
        
    Next outerKey
    
    ws.Range("a" & rowPtr).value = "Clause 11"
    rowPtr = rowPtr + 1
    
    For Each outerKey In curentUpAsWriteFormatDict("upClause11").keys
    
        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause11")(outerKey), True, True, True
        rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause11")(outerKey).Count + 1
        
    Next outerKey
    
    ws.Range("a" & rowPtr).value = "Clause 12a"
    rowPtr = rowPtr + 1
    
    For Each outerKey In curentUpAsWriteFormatDict("upClause12a").keys
    
        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause12a")(outerKey), True, True, True
        rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause12a")(outerKey).Count + 1
        
    Next outerKey
    
    ws.Range("a" & rowPtr).value = "Clause 12b Fabrics"
    rowPtr = rowPtr + 1
    
    For Each outerKey In curentUpAsWriteFormatDict("upClause12bFabrics").keys
    
        Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause12bFabrics")(outerKey), True, True, True
        rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause12bFabrics")(outerKey).Count
        
    Next outerKey
    
    rowPtr = rowPtr + 1
    
    ws.Range("a" & rowPtr).value = "Clause 12b Garments"
    rowPtr = rowPtr + 1
    
    Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause12bGarments"), True, True, True
    rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause12bGarments").Count + 1
    
    ws.Range("a" & rowPtr).value = "Clause 13"
    rowPtr = rowPtr + 1
    
    Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause13"), True, True, True
    rowPtr = rowPtr + curentUpAsWriteFormatDict("upClause13").Count + 1
        
    ws.Range("a" & rowPtr).value = "Clause 14"
    rowPtr = rowPtr + 1
    
    Application.Run "dictionary_utility_functions.PutDictionaryValuesIntoWorksheet", ws.Range("a" & rowPtr), curentUpAsWriteFormatDict("upClause14"), True, True, True

End Sub

Sub upSummaryTemplete()

    'just Value & qty. templete, to be modify as need
    Dim allUpDic As Object
    Dim jsonPathArr As Variant

    jsonPathArr = Application.Run("general_utility_functions.returnSelectedFilesFullPathArr", "D:\Temp\UP Draft\Draft 2024\json-all-up-clause")  ' JSON file path

    If Not UBound(jsonPathArr) = 1 Then
        MsgBox "Please select only one JSON file"
        Exit Sub
    End If
    
    Set allUpDic = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", jsonPathArr(1))
    
    Dim outerKey, innerKey1, innerKey2 As Variant
    Dim sumOfFabricsQty, sumOfExportValue As Variant

    sumOfFabricsQty = 0
    sumOfExportValue = 0

    For Each outerKey In allUpDic.keys

        For Each innerKey1 In allUpDic(outerKey)("upClause7").keys
        
            If allUpDic(outerKey)("upClause7")(innerKey1)("isLcValueExistInEuro") Then
            
                sumOfExportValue = sumOfExportValue + allUpDic(outerKey)("upClause7")(innerKey1)("lcValueInEuro")
            Else
                sumOfExportValue = sumOfExportValue + allUpDic(outerKey)("upClause7")(innerKey1)("lcValueInUsd")
            
            End If
            
            If allUpDic(outerKey)("upClause7")(innerKey1)("isFabQtyExistInMtr") Then
            
                sumOfFabricsQty = sumOfFabricsQty + allUpDic(outerKey)("upClause7")(innerKey1)("fabricsQtyInMtr")
            Else
                sumOfFabricsQty = sumOfFabricsQty + allUpDic(outerKey)("upClause7")(innerKey1)("fabricsQtyInYds")
            
            End If
            
        Next innerKey1

    Next outerKey
    
    Debug.Print sumOfExportValue, sumOfFabricsQty
    
End Sub
