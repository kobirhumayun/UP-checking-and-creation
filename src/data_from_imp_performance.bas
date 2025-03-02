Attribute VB_Name = "data_from_imp_performance"
Option Explicit

Private Function createUseGroupDic() As Object

    Dim useGroupDict As Object ' use group be changed as requirement
    Set useGroupDict = CreateObject("Scripting.Dictionary") ' UP raw materials group against import performance raw materials

    Dim isYarn As Object ' use group be changed as requirement
    Set isYarn = CreateObject("Scripting.Dictionary") ' UP raw materials group against import performance raw materials
    'alternatively UP raw materials assign import performance raw materials id

    Dim yarnUseGroupDict As Object ' use group be changed as requirement
    Set yarnUseGroupDict = CreateObject("Scripting.Dictionary") ' UP raw materials group against import performance raw materials
    'alternatively UP raw materials assign import performance raw materials id

    Dim nonYarnUseGroupDict As Object ' use group be changed as requirement
    Set nonYarnUseGroupDict = CreateObject("Scripting.Dictionary") ' UP raw materials group against import performance raw materials
    'alternatively UP raw materials assign import performance raw materials id

    Set yarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", yarnUseGroupDict, "cotton", _
    Array("50% Cotton 50% Modal Yarn", _
    "Carded Cotton Yarn", _
    "Cotton Carded Lycra Yarn", _
    "Cotton Carded Yarn", _
    "Cotton / Lycra Yarn", _
    "Cotton / Polyester Yarn", _
    "Cotton Yarn", _
    "COTTON YARN", _
    "LYOCELL YARN", _
    "TENCEL YARN", _
    "Hemp Yarn", _
    "True Hemp Yarn"))

    Set yarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", yarnUseGroupDict, "polyester", _
    Array("35% Cotton 65% Polyester Yarn", _
    "65% Polyester 35% Cotton Spandex Yarn", _
    "65% Polyester 35% Rayon Yarn", _
    "65% Tencel 35% Cotton Yarn", _
    "Polyester Yarn", _
    "65% Polyester 35% Cotton Yarn", _
    "Viscose Rayon Yarn"))

    Set yarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", yarnUseGroupDict, "spandex", _
    Array("Lycra Yarn", _
    "Spandex Bare Yarn", _
    "Spandex Yarn"))

    useGroupDict.Add "yarnUseGroupDict", yarnUseGroupDict

    Set isYarn = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", isYarn, "yarn", yarnUseGroupDict.keys) 'dynamic array elements
    useGroupDict.Add "isYarn", isYarn

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Desizing Agent / Enzyme", _
    Array("Desizing Agent", "Enzyme"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Acetic Acid", _
    Array("Acetic Acid"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Binder", _
    Array("Binder"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Caustic Soda", _
    Array("Caustic Soda"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Detergent", _
    Array("Detergent"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Dispersing Agent", _
    Array("Dispersing Agent"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Fixing Agent", _
    Array("Fixing Agent"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Hydrogen Peroxide", _
    Array("Hydrogen Peroxide"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Modified Starch", _
    Array("Modified Starches"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "PVA", _
    Array("PVA"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Reducing Agent", _
    Array("Reducing Agent"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Sequestering Agent", _
    Array("Sequestering Agent"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Sodium Hydro Sulphate", _
    Array("Sodium Hydro Sulphite"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Softener", _
    Array("Softening Agent (Softener)"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Stabilizing Agent(Estabilizador FE)", _
    Array("Stabilizing Agent"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Stretch Wrapping Film", _
    Array("Stretch Wrapping Film"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Sulphur Dyes", _
    Array("Sulphur Dyes"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Sulphuric Acid", _
    Array("Sulphuric Acid"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Vat Dyes", _
    Array("Vat Dyes"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Vat Dyes (Indigo Granular)", _
    Array("Vat Dyes (Indigo Granular)"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Water Decoloring Agent", _
    Array("Water Decoloring Agent"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Sodium Hypochloride", _
    Array("Sodium Hypochloride"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Wax", _
    Array("Waxes"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Resin", _
    Array("Resin"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Wetting Agent", _
    Array("Wetting Agent", "Mercerizing Agent (Wetting Agent)"))
    
    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Pumice Stone", _
    Array("Pumice Stone"))
    
    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Natural Garnet", _
    Array("Natural Garnet"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Hydroxylamine", _
    Array("Hydroxylamine"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Bleaching Powder", _
    Array("Bleaching Powder"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Finishing Agent", _
    Array("Finishing Agent"))

    Set nonYarnUseGroupDict = Application.Run("dictionary_utility_functions.AddKeysWithPrimary", nonYarnUseGroupDict, "Decision Pending Group", _
    Array("Activated Carbon", "Polymers"))

    useGroupDict.Add "nonYarnUseGroupDict", nonYarnUseGroupDict

    Set createUseGroupDic = useGroupDict

End Function


Private Function classifiedDbDicFromImpPerformance(importPerformanceFilePath As String) As Object
    'this function return as required classified Db Dic from import performance

    Dim returnDic As Object
    Set returnDic = CreateObject("Scripting.Dictionary")

    Dim yarnClassifiedDbDic As Object
    Set yarnClassifiedDbDic = CreateObject("Scripting.Dictionary")

    Dim yarnGroupNameDic As Object
    Set yarnGroupNameDic = CreateObject("Scripting.Dictionary")

    Dim CottonYarnLocalOrImpClassifiedDbDic As Object
    Set CottonYarnLocalOrImpClassifiedDbDic = CreateObject("Scripting.Dictionary")

    CottonYarnLocalOrImpClassifiedDbDic.Add "importCtnAsBillOfEntry", CreateObject("Scripting.Dictionary") ' import cotton
    CottonYarnLocalOrImpClassifiedDbDic.Add "localCtnAsLc", CreateObject("Scripting.Dictionary") ' local cotton

    Dim nonYarnClassifiedDbDic As Object
    Set nonYarnClassifiedDbDic = CreateObject("Scripting.Dictionary")

    Dim notDefUseGroupDic As Object
    Set notDefUseGroupDic = CreateObject("Scripting.Dictionary")

    Dim useGroupDict As Object
    Set useGroupDict = Application.Run("data_from_imp_performance.createUseGroupDic")

    Dim isYarn As Object ' use group be changed as requirement
    Dim yarnUseGroupDict As Object ' use group be changed as requirement
    Dim nonYarnUseGroupDict As Object ' use group be changed as requirement

    Set isYarn = useGroupDict("isYarn")
    Set yarnUseGroupDict = useGroupDict("yarnUseGroupDict")
    Set nonYarnUseGroupDict = useGroupDict("nonYarnUseGroupDict")


'    importPerformanceFilePath = ActiveWorkbook.path & Application.PathSeparator & "Import Performance Statement of PDL-2025-2026.xlsx" ' file name will be change after change period

    Dim impBillAndMushakDb As Object
    Set impBillAndMushakDb = Application.Run("utilityFunction.CombinedAllSheetsMushakOrBillOfEntryDbDict", importPerformanceFilePath)

    Application.ScreenUpdating = False
    Dim importPerformanceWb As Workbook
    Set importPerformanceWb = Workbooks.Open(importPerformanceFilePath)
    
    Dim yarnImportWs As Worksheet
    Set yarnImportWs = importPerformanceWb.Worksheets("Yarn (Import)")

    Dim garmentsYarnBillOfEntry As Object
    Set garmentsYarnBillOfEntry = Application.Run("utilityFunction.importPerformanceCommentedBillOfEntryOrMushakDbFromProvidedSheet", yarnImportWs, 4, 3, 7, 8)

    importPerformanceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
    Dim dicKey As Variant
    
    For Each dicKey In garmentsYarnBillOfEntry.keys
        
        If Application.Run("general_utility_functions.isStrPatternExist", garmentsYarnBillOfEntry(dicKey)("comment"), "garments", True, True, True) Then
                'Remove garments yarn bill of entry
            If impBillAndMushakDb.Exists(dicKey) Then
            
                impBillAndMushakDb.Remove dicKey
                
            Else
                
                MsgBox "Garments bill of entry " & dicKey & " not exitst in import performance DB dictionary"
                Exit Function
                
            End If
        End If
    
    Next dicKey

    Dim tempDic As Object
    Dim extractLc As String

    Dim removedAllInvalidChrFromImpRawMaterialsDes As Variant
    Dim removedAllInvalidChrFromPreDefclassifiedDbDic As Variant

    For Each dicKey In impBillAndMushakDb.keys

        removedAllInvalidChrFromImpRawMaterialsDes = Application.Run("general_utility_functions.RemoveInvalidChars", impBillAndMushakDb(dicKey)("Description"))   'remove all invalid characters


        If isYarn.Exists(removedAllInvalidChrFromImpRawMaterialsDes) Then


            If yarnUseGroupDict.Exists(removedAllInvalidChrFromImpRawMaterialsDes) Then 'pick use group raw materials name from pre-defiened yarn use group dictionary

                removedAllInvalidChrFromPreDefclassifiedDbDic = Application.Run("general_utility_functions.RemoveInvalidChars", yarnUseGroupDict(removedAllInvalidChrFromImpRawMaterialsDes))   'remove all invalid characters

            Else

                removedAllInvalidChrFromPreDefclassifiedDbDic = "notDefUseGroup"

                notDefUseGroupDic(removedAllInvalidChrFromImpRawMaterialsDes) = impBillAndMushakDb(dicKey)("Description") 'for update pre-defiened use group dictionary

            End If


            If Not yarnClassifiedDbDic.Exists(removedAllInvalidChrFromPreDefclassifiedDbDic) Then ' create classified dictionary for return

                yarnClassifiedDbDic.Add removedAllInvalidChrFromPreDefclassifiedDbDic, CreateObject("Scripting.Dictionary")

            End If

            If impBillAndMushakDb(dicKey)("UsedQty") = 0 Then

                yarnClassifiedDbDic(removedAllInvalidChrFromPreDefclassifiedDbDic).Add dicKey, impBillAndMushakDb(dicKey)

            End If

            yarnGroupNameDic(dicKey) = removedAllInvalidChrFromPreDefclassifiedDbDic


            If removedAllInvalidChrFromPreDefclassifiedDbDic = "cotton" Then
            ' import & local cotton yarn devided, import yarn no group but local yarn grouped by lc

                If Left$(impBillAndMushakDb(dicKey)("BillOfEntryOrMushak"), 2) = "C-" Then

                    If impBillAndMushakDb(dicKey)("UsedQty") = 0 Then

                        CottonYarnLocalOrImpClassifiedDbDic("importCtnAsBillOfEntry").Add dicKey, impBillAndMushakDb(dicKey) ' just add, no group

                    End If

                Else

                    extractLc = Left$(impBillAndMushakDb(dicKey)("LC"), Len(impBillAndMushakDb(dicKey)("LC")) - 11)

                    If Not CottonYarnLocalOrImpClassifiedDbDic("localCtnAsLc").Exists(extractLc) Then ' create classified dictionary for return

                        CottonYarnLocalOrImpClassifiedDbDic("localCtnAsLc").Add extractLc, CreateObject("Scripting.Dictionary")

                    End If

                    If impBillAndMushakDb(dicKey)("UsedQty") = 0 Then

                        CottonYarnLocalOrImpClassifiedDbDic("localCtnAsLc")(extractLc).Add dicKey, impBillAndMushakDb(dicKey) ' grouped by lc then add mushak

                    End If

                End If

            End If

        Else

            If nonYarnUseGroupDict.Exists(removedAllInvalidChrFromImpRawMaterialsDes) Then 'pick use group raw materials name from pre-defiened non yarn use group dictionary

                removedAllInvalidChrFromPreDefclassifiedDbDic = Application.Run("general_utility_functions.RemoveInvalidChars", nonYarnUseGroupDict(removedAllInvalidChrFromImpRawMaterialsDes))   'remove all invalid characters

            Else

                removedAllInvalidChrFromPreDefclassifiedDbDic = "notDefUseGroup"

                notDefUseGroupDic(removedAllInvalidChrFromImpRawMaterialsDes) = impBillAndMushakDb(dicKey)("Description") 'for update pre-defiened use group dictionary

            End If


            If Not nonYarnClassifiedDbDic.Exists(removedAllInvalidChrFromPreDefclassifiedDbDic) Then ' create classified dictionary for return

                nonYarnClassifiedDbDic.Add removedAllInvalidChrFromPreDefclassifiedDbDic, CreateObject("Scripting.Dictionary")

            End If

            If impBillAndMushakDb(dicKey)("UsedQty") = 0 Then

                nonYarnClassifiedDbDic(removedAllInvalidChrFromPreDefclassifiedDbDic).Add dicKey, impBillAndMushakDb(dicKey)

            End If

        End If

    Next


    If notDefUseGroupDic.Count > 0 Then ' only for update pre-defiened use group dictionary

        For Each dicKey In notDefUseGroupDic.keys
            ' Debug.Print notDefUseGroupDic(dicKey) 'for copy to use group
            ' Debug.Print dicKey 'for copy to use group
        Next dicKey

        MsgBox "some raw materials not defined in use group"
        Exit Function

    End If

    returnDic.Add "yarnClassifiedDbDic", yarnClassifiedDbDic
    returnDic.Add "yarnGroupNameDic", yarnGroupNameDic
    returnDic.Add "CottonYarnLocalOrImpClassifiedDbDic", CottonYarnLocalOrImpClassifiedDbDic
    returnDic.Add "nonYarnClassifiedDbDic", nonYarnClassifiedDbDic

    Set classifiedDbDicFromImpPerformance = returnDic

End Function

