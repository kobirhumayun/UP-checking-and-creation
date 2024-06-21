Attribute VB_Name = "dedo_consumption"
Option Explicit


Private Function ropeDenimFabricsDyedBlack() As Object
    'this function return a dictionary of dedo consumption rate for Rope Denim Fabrics (Dyed) Black.

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Caustic soda (NaOH) Solid_Sl_1", _
    "2. Wetting agent and Detergent _Sl_2", _
    "3. Sequestering agent_Sl_3", _
    "4.  Fixing Agent_Sl_4", _
    "5. Sulphur black (Powder)_Sl_5", _
    "Or Sulphur black  (Liquid), Concentration active substance not more than 25%_Sl_6", _
    "6. Reducing agent DP_Sl_7", _
    "7. Acetic Acid_Sl_8", _
    "8. Rope opening Chemical_Sl_9", _
    "9. Hydrogent Peroxide_Sl_10", _
    "1. Starch /Modified Starch_Sl_11", _
    "2. Binder_Sl_12", _
    "3. Wax_Sl_13", _
    "4. PVA_Sl_14", _
    "1. Desizing agent_Sl_15", _
    "2. Wetting agent and Detergent_Sl_16", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_19", _
    "2. Acetic acid_Sl_20" _
    )

    propertiesValArr = Array( _
    6.7, _
    0.6, _
    0.3, _
    0.1, _
    2.5, _
    10, _
    2.9, _
    0.62, _
    1.25, _
    0.5, _
    6.5, _
    1, _
    0.5, _
    0.3, _
    0.15, _
    0.15, _
    1, _
    0.18 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set ropeDenimFabricsDyedBlack = dedoConRateDic

End Function

Private Function ropeDenimFabricsDyedBlackMercerization() As Object
    'this function return a dictionary of dedo consumption rate for Rope Denim Fabrics (Dyed) Black. Mercerization

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Caustic soda (Solid)_Sl_17", _
    "2. Acetic acid_Sl_18" _
    )

    propertiesValArr = Array( _
    17, _
    0.15 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set ropeDenimFabricsDyedBlackMercerization = dedoConRateDic

End Function


Private Function ropeDenimFabricsDyedIndigo() As Object
    'this function return a dictionary of dedo consumption rate for Rope Denim Fabrics (Dyed) Indigo

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Caustic soda (NaOH)_Sl_21", _
    "2. Wetting agent and Detergent_Sl_22", _
    "3. Sequestering agent_Sl_23", _
    "4. Fixing Agent_Sl_24", _
    "5. Vat Dyes (Powder/Solid)_Sl_25", _
    "Or Vat Dyes (Liquid), Concentration active substance not more than  30%_Sl_26", _
    "6. Sodium Hydro Sulphite_Sl_27", _
    "7. Rope opening Chemical_Sl_28", _
    "8. Acetic acid_Sl_29", _
    "9. Dispersing Agent_Sl_30", _
    "1. Starch /Modified Starch_Sl_31", _
    "2. Binder_Sl_32", _
    "3. Wax_Sl_33", _
    "4. PVA_Sl_34", _
    "1. Desizing agent_Sl_35", _
    "2. Wetting agent and Detergent_Sl_36", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_39", _
    "2. Acetic acid_Sl_40" _
    )

    propertiesValArr = Array( _
    6.7, _
    0.6, _
    0.3, _
    0.1, _
    2, _
    6.67, _
    4, _
    1.25, _
    0.3, _
    0.1, _
    6.5, _
    1, _
    0.5, _
    0.3, _
    0.15, _
    0.15, _
    1, _
    0.18 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set ropeDenimFabricsDyedIndigo = dedoConRateDic

End Function

Private Function ropeDenimFabricsDyedIndigoMercerization() As Object
    'this function return a dictionary of dedo consumption rate for Rope Denim Fabrics (Dyed) Indigo Mercerization

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Caustic soda (Solid)_Sl_37", _
    "2. Acetic acid_Sl_38" _
    )

    propertiesValArr = Array( _
    17, _
    0.15 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set ropeDenimFabricsDyedIndigoMercerization = dedoConRateDic

End Function

Private Function ropeDenimFabricsDyed() As Object
    'this function return a dictionary of dedo consumption rate for Rope Denim Fabrics (Dyed)

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Caustic soda (NaOH)_Sl_41", _
    "2. Wetting agent and Detergent_Sl_42", _
    "3. Sequestering agent_Sl_43", _
    "4. Fixing Agent_Sl_44", _
    "5. Vat Dyes (Powder/Solid)_Sl_45", _
    "Or Vat Dyes  (Liquid), Concentration active substance not more than 30%_Sl_46", _
    "6. Sulphur black (Powder)_Sl_47", _
    "Or Sulphur back  (Liquid), Concentration active substance not more than 30%_Sl_48", _
    "7. Rope openninng chemical_Sl_49", _
    "8. Reducing agent DP_Sl_50", _
    "9. Sodium Hydro Sulphite_Sl_51", _
    "10. Acetic acid_Sl_52", _
    "11. Dispersing Agent/Leveling Agent_Sl_53", _
    "1. Starch /Modified Starch_Sl_54", _
    "2. Binder_Sl_55", _
    "3. Wax_Sl_56", _
    "4. PVA_Sl_57", _
    "1. Desizing agent_Sl_58", _
    "2. Wetting agent and Detergent_Sl_59", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_62", _
    "2. Acetic acid_Sl_63" _
    )

    propertiesValArr = Array( _
    6.7, _
    0.6, _
    0.3, _
    0.1, _
    2, _
    6.67, _
    2.5, _
    10, _
    1.25, _
    2.1, _
    4.5, _
    0.25, _
    0.1, _
    6.5, _
    1, _
    0.5, _
    0.3, _
    0.15, _
    0.15, _
    1, _
    0.18 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set ropeDenimFabricsDyed = dedoConRateDic

End Function


Private Function ropeDenimFabricsDyedMercerization() As Object
    'this function return a dictionary of dedo consumption rate for Rope Denim Fabrics (Dyed) Mercerization

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Caustic soda (Solid)_Sl_60", _
    "2. Acetic acid_Sl_61" _
    )

    propertiesValArr = Array( _
    17, _
    0.15 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set ropeDenimFabricsDyedMercerization = dedoConRateDic

End Function

Private Function denimFabricsOverDyedSolidDyed() As Object
    'this function return a dictionary of dedo consumption rate for Denim Fabrics Over Dyed / Solid Dyed.

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Starch /Modified Starch_Sl_64", _
    "2. Binder_Sl_65", _
    "3. Wax_Sl_66", _
    "4. PVA_Sl_67", _
    "1. Caustic soda (Liquid) , (Concentration active substance not more than 60%)_Sl_70", _
    "2. Wetting agent and Detergent_Sl_71", _
    "3. Sequestering agent_Sl_72", _
    "4. Sulphur black (Liquid), (Concentration active substance not more than 30%)_Sl_73", _
    "5. Reducing agent DP_Sl_74", _
    "6. Acetic acid_Sl_75", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_76", _
    "2. Acetic acid_Sl_77" _
    )

    propertiesValArr = Array( _
    6.5, _
    1, _
    0.4, _
    0.3, _
    8, _
    0.65, _
    0.3, _
    10, _
    3, _
    0.15, _
    1, _
    0.18 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set denimFabricsOverDyedSolidDyed = dedoConRateDic

End Function

Private Function denimFabricsOverDyedSolidDyedMercerization() As Object
    'this function return a dictionary of dedo consumption rate for Denim Fabrics Over Dyed / Solid Dyed. Mercerization

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Caustic soda (Solid)_Sl_68", _
    "2. Acetic acid_Sl_69" _
    )

    propertiesValArr = Array( _
    17, _
    0.15 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set denimFabricsOverDyedSolidDyedMercerization = dedoConRateDic

End Function


Private Function denimFabricsCoatedandPigment() As Object
    'this function return a dictionary of dedo consumption rate for Denim Fabrics (Coated and Pigment)

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Resin_Sl_78", _
    "2. Binder_Sl_79", _
    "3. Coating Pigment_Sl_80", _
    "4. Wetting Agent and Detergent_Sl_81", _
    "5. Softner_Sl_82", _
    "6. Thickener_Sl_83", _
    "7. Wrinkle resistant agent (Optional)_Sl_84", _
    "8. Deniart Trinity (Special Prodcut for only Effect)_Sl_85" _
    )

    propertiesValArr = Array( _
    10, _
    6, _
    1, _
    0.5, _
    1, _
    3, _
    0.2, _
    0.1 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set denimFabricsCoatedandPigment = dedoConRateDic

End Function


Private Function denimFabricsPFDFinished() As Object
    'this function return a dictionary of dedo consumption rate for Denim Fabrics (PFD) Finished.

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Starch /Modified Starch_Sl_86", _
    "2. Binder_Sl_87", _
    "3. Wax_Sl_88", _
    "4. PVA_Sl_89", _
    "1. Hydrogent Peroxide (liquid), Concentration 45%_Sl_90", _
    "2. Sequestering agent_Sl_91", _
    "3. Caustic Soda (Solid)_Sl_92", _
    "Or Caustic soda (Liquid) Local , (Concentration active substance not more than 60%)_Sl_93", _
    "4. Peroxide Stabilizing Agent_Sl_94", _
    "5. Wetting agent and Detergent_Sl_95", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_96", _
    "2. Acetic acid_Sl_97" _
    )

    propertiesValArr = Array( _
    6.5, _
    1, _
    0.5, _
    0.3, _
    16, _
    0.37, _
    4, _
    10, _
    1.73, _
    0.7, _
    1, _
    0.18 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set denimFabricsPFDFinished = dedoConRateDic

End Function


Private Function denimFabricsECRUFinished() As Object
    'this function return a dictionary of dedo consumption rate for Denim Fabrics (ECRU) Finished.

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1. Starch /Modified Starch_Sl_98", _
    "2. Binder_Sl_99", _
    "3. Wax_Sl_100", _
    "4. PVA_Sl_101", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_102", _
    "2. Acetic acid_Sl_103" _
    )

    propertiesValArr = Array( _
    6.5, _
    1, _
    0.4, _
    0.3, _
    1, _
    0.18 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set denimFabricsECRUFinished = dedoConRateDic

End Function


Private Function denimFabricDyedEtpWtp() As Object
    'this function return a dictionary of dedo consumption rate for Denim Fabric (Dyed) ETP & WTP

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "1.98% Sulfuric Acid (Commercial)_Sl_104", _
    "2. Water decoloring Agent_Sl_105", _
    "1. Sodium Hydroxide (NaOH)_Sl_106", _
    "2. Cataionic poly electrolyte_Sl_107", _
    "3. Alum_Sl_108", _
    "4. Sodium  Hypochloride ( NaOCI)_Sl_109" _
    )

    propertiesValArr = Array( _
    7, _
    0.5, _
    0.1871, _
    0.0003, _
    0.0972, _
    0.0002 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set denimFabricDyedEtpWtp = dedoConRateDic

End Function


Private Function denimFabricPacking() As Object
    'this function return a dictionary of dedo consumption rate for Denim Fabric (Packing)

    Dim dedoConRateDic As Object

    Dim propertiesArr, propertiesValArr As Variant

    propertiesArr = Array( _
    "Strech Wrapping Flim_Sl_110", _
    "Paper Tube_Sl_111" _
    )

    propertiesValArr = Array( _
    0.16, _
    0.8 _
    )

    Set dedoConRateDic = Application.Run("dictionary_utility_functions.CreateDicWithProvidedKeysAndValues", propertiesArr, propertiesValArr)

    Set denimFabricPacking = dedoConRateDic

End Function


Private Function dedoConRateToActualQtyCalculation(dedoConRateDic As Object, qty As Variant) As Object
    'this function received dedo consumption rate dictionary and oparation Qty. then calculate the actual raw materials Qty.
    'and return calculated new dictionary

    ' Create new dictionary
    Dim dictNew As Object
    Set dictNew = CreateObject("Scripting.Dictionary")

    Dim dictKey As Variant

    For Each dictKey In dedoConRateDic.keys

        dictNew.Add dictKey, dedoConRateDic(dictKey) / 100 * qty

    Next

    Set dedoConRateToActualQtyCalculation = dictNew

End Function


Private Function combineAllDedoConDicAfterCalculateActualQty( _
        ropeDenimFabricsDyedBlackQty As Variant, _
        ropeDenimFabricsDyedBlackMercerizationQty As Variant, _
        ropeDenimFabricsDyedIndigoQty As Variant, _
        ropeDenimFabricsDyedIndigoMercerizationQty As Variant, _
        ropeDenimFabricsDyedQty As Variant, _
        ropeDenimFabricsDyedMercerizationQty As Variant, _
        denimFabricsOverDyedSolidDyedQty As Variant, _
        denimFabricsOverDyedSolidDyedMercerizationQty As Variant, _
        denimFabricsCoatedandPigmentQty As Variant, _
        denimFabricsPFDFinishedQty As Variant, _
        denimFabricsECRUFinishedQty As Variant, _
        denimFabricDyedEtpWtpQty As Variant, _
        denimFabricPackingQty As Variant) As Object
    'this function received all process qty. and return a combined dictionary, after calculate actual raw material qty. by dedo consumption rate

    Dim combinedDic As Object
    Set combinedDic = CreateObject("Scripting.Dictionary")

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.ropeDenimFabricsDyedBlack"), ropeDenimFabricsDyedBlackQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.ropeDenimFabricsDyedBlackMercerization"), ropeDenimFabricsDyedBlackMercerizationQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.ropeDenimFabricsDyedIndigo"), ropeDenimFabricsDyedIndigoQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.ropeDenimFabricsDyedIndigoMercerization"), ropeDenimFabricsDyedIndigoMercerizationQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.ropeDenimFabricsDyed"), ropeDenimFabricsDyedQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.ropeDenimFabricsDyedMercerization"), ropeDenimFabricsDyedMercerizationQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.denimFabricsOverDyedSolidDyed"), denimFabricsOverDyedSolidDyedQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.denimFabricsOverDyedSolidDyedMercerization"), denimFabricsOverDyedSolidDyedMercerizationQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.denimFabricsCoatedandPigment"), denimFabricsCoatedandPigmentQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.denimFabricsPFDFinished"), denimFabricsPFDFinishedQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.denimFabricsECRUFinished"), denimFabricsECRUFinishedQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.denimFabricDyedEtpWtp"), denimFabricDyedEtpWtpQty))

    Set combinedDic = Application.Run("dictionary_utility_functions.mergeDict", combinedDic, _
    Application.Run("dedo_consumption.dedoConRateToActualQtyCalculation", Application.Run("dedo_consumption.denimFabricPacking"), denimFabricPackingQty))


    Set combineAllDedoConDicAfterCalculateActualQty = combinedDic

End Function


Private Function appliedUsedPercentageSpecificRawMaterials(combineDicAfterCalculateActualQty As Object, usedPercentageSpecificRawMaterialsDict As Object) As Object
    'this function received combined dictionary after calculate actual Qty. and specific raw materials used percentage dict then
    ' specific raw materials Qty. set as provided percentage and return modified dictionary

    Dim dictKey  As Variant

    Dim msgStr As String
    msgStr = ""

    For Each dictKey In usedPercentageSpecificRawMaterialsDict.keys

        If usedPercentageSpecificRawMaterialsDict(dictKey) <> 100 Then ' for throw msg against keys only <> 100 percentage

            If combineDicAfterCalculateActualQty.Exists(dictKey) Then

                combineDicAfterCalculateActualQty(dictKey) = combineDicAfterCalculateActualQty(dictKey) * usedPercentageSpecificRawMaterialsDict(dictKey) / 100 ' calculate & assign percentage

                msgStr = msgStr & dictKey & " = " & usedPercentageSpecificRawMaterialsDict(dictKey) & "%" & Chr(10)

            Else

                MsgBox "Dictionary Key """ & dictKey & """ Not Found"
                Exit Function

            End If

        End If

    Next

    If msgStr = "" Then

        MsgBox "All raw materials using 100%", , "Raw materials use as DEDO"

    Else

        MsgBox msgStr, , "Raw materials use as DEDO"

    End If

    Set appliedUsedPercentageSpecificRawMaterials = combineDicAfterCalculateActualQty

End Function

Private Function combineDicSaveAsJsonForUsedPercentageSpecificRawMaterials(combineDic As Object, jsonPath As String)
    'combined dictionary save as json to take actual properties name
    'initially all properties value set 100, as requirements modify json file manually

    Dim dictKey As Variant

    For Each dictKey In combineDic.keys

        combineDic(dictKey) = 100

    Next dictKey

    Application.Run "JsonUtilityFunction.SaveDictionaryToJsonTextFile", combineDic, jsonPath & Application.PathSeparator & _
    "used-percentage-specific-raw-materials" & ".json"

End Function


Private Function sumRawMaterialsAsGroupAndAddToDic(qtyAsGroupDic As Object, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials As Object, _
    rawMaterialsGroupName As Variant, rawMaterialsGroupArr As Variant) As Object
    ' this function received Qty. group dic, final raw materials caculated dic, raw materials group name and raw metatials group arr,
    ' then sum all raw metarials Qty. and add to Qty. group dic
    Dim tempSum As Variant

    tempSum = Application.Run("dictionary_utility_functions.sumOfProvidedKeys", allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, rawMaterialsGroupArr)

   Set qtyAsGroupDic = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", qtyAsGroupDic, rawMaterialsGroupName, tempSum)

    ' Return the modified qtyAsGroupDic
    Set sumRawMaterialsAsGroupAndAddToDic = qtyAsGroupDic

End Function


Private Function finalRawMaterialsQtyCalculatedAsGroup( _
    ropeDenimFabricsDyedBlackQty As Variant, _
    ropeDenimFabricsDyedBlackMercerizationQty As Variant, _
    ropeDenimFabricsDyedIndigoQty As Variant, _
    ropeDenimFabricsDyedIndigoMercerizationQty As Variant, _
    ropeDenimFabricsDyedQty As Variant, _
    ropeDenimFabricsDyedMercerizationQty As Variant, _
    denimFabricsOverDyedSolidDyedQty As Variant, _
    denimFabricsOverDyedSolidDyedMercerizationQty As Variant, _
    denimFabricsCoatedandPigmentQty As Variant, _
    denimFabricsPFDFinishedQty As Variant, _
    denimFabricsECRUFinishedQty As Variant, _
    denimFabricDyedEtpWtpQty As Variant, _
    denimFabricPackingQty As Variant) As Object
    'this function received all process qty. and return a dictionary, after calculate actual raw materials qty.'s sum as raw materials group
    'note only this function call from outside

    Dim allDedoConDicAfterCalculateActualQty As Object
    Dim allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials As Object

    Dim usedPercentageSpecificRawMaterialsDict As Object
    ' Set usedPercentageSpecificRawMaterialsDict = CreateObject("Scripting.Dictionary")

    ' Set usedPercentageSpecificRawMaterialsDict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", usedPercentageSpecificRawMaterialsDict, "5. Sulphur black (Powder)_Sl_5", 0)
    ' Set usedPercentageSpecificRawMaterialsDict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", usedPercentageSpecificRawMaterialsDict, "5. Vat Dyes (Powder/Solid)_Sl_25", 0)
    ' Set usedPercentageSpecificRawMaterialsDict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", usedPercentageSpecificRawMaterialsDict, "5. Vat Dyes (Powder/Solid)_Sl_45", 0)
    ' Set usedPercentageSpecificRawMaterialsDict = Application.Run("dictionary_utility_functions.addKeysAndValueToDic", usedPercentageSpecificRawMaterialsDict, "6. Sulphur black (Powder)_Sl_47", 0)


    Set allDedoConDicAfterCalculateActualQty = Application.Run("dedo_consumption.combineAllDedoConDicAfterCalculateActualQty", _
    ropeDenimFabricsDyedBlackQty, _
    ropeDenimFabricsDyedBlackMercerizationQty, _
    ropeDenimFabricsDyedIndigoQty, _
    ropeDenimFabricsDyedIndigoMercerizationQty, _
    ropeDenimFabricsDyedQty, _
    ropeDenimFabricsDyedMercerizationQty, _
    denimFabricsOverDyedSolidDyedQty, _
    denimFabricsOverDyedSolidDyedMercerizationQty, _
    denimFabricsCoatedandPigmentQty, _
    denimFabricsPFDFinishedQty, _
    denimFabricsECRUFinishedQty, _
    denimFabricDyedEtpWtpQty, _
    denimFabricPackingQty)

    Dim jsonPath As String
    jsonPath = ActiveWorkbook.path & Application.PathSeparator & "json-used-percentage-specific-raw-materials"

        'uncomment just for save first time, then again comment bellow function call
        'modify json file as requirements
    ' Application.Run "dedo_consumption.combineDicSaveAsJsonForUsedPercentageSpecificRawMaterials", allDedoConDicAfterCalculateActualQty, jsonPath

    Set usedPercentageSpecificRawMaterialsDict = Application.Run("JsonUtilityFunction.LoadDictionaryFromJsonTextFile", jsonPath & Application.PathSeparator & "used-percentage-specific-raw-materials" & ".json")

    Set allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials = Application.Run("dedo_consumption.appliedUsedPercentageSpecificRawMaterials", allDedoConDicAfterCalculateActualQty, _
    usedPercentageSpecificRawMaterialsDict)

    Dim qtyAsGroupDic As Object
    Set qtyAsGroupDic = CreateObject("Scripting.Dictionary")

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Sulphur Dyes", _
    Array( _
    "Or Sulphur black  (Liquid), Concentration active substance not more than 25%_Sl_6", _
    "Or Sulphur back  (Liquid), Concentration active substance not more than 30%_Sl_48", _
    "4. Sulphur black (Liquid), (Concentration active substance not more than 30%)_Sl_73" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Vat Dyes", _
    Array( _
    "Or Vat Dyes (Liquid), Concentration active substance not more than  30%_Sl_26", _
    "Or Vat Dyes  (Liquid), Concentration active substance not more than 30%_Sl_46" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Reducing Agent", _
    Array( _
    "6. Reducing agent DP_Sl_7", _
    "8. Reducing agent DP_Sl_50", _
    "5. Reducing agent DP_Sl_74" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Wetting Agent", _
    Array( _
    "2. Wetting agent and Detergent _Sl_2", _
    "2. Wetting agent and Detergent_Sl_16", _
    "2. Wetting agent and Detergent_Sl_22", _
    "2. Wetting agent and Detergent_Sl_36", _
    "2. Wetting agent and Detergent_Sl_42", _
    "2. Wetting agent and Detergent_Sl_59", _
    "2. Wetting agent and Detergent_Sl_71", _
    "5. Wetting agent and Detergent_Sl_95" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Softener", _
    Array( _
    "8. Rope opening Chemical_Sl_9", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_19", _
    "7. Rope opening Chemical_Sl_28", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_39", _
    "7. Rope openninng chemical_Sl_49", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_62", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_76", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_96", _
    "1. Softening  agent : (Cationic, Anionic, Non ionic) (Under different Trade/Code/Chemical names of different manufacturers/Suppliers)_Sl_102" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Acetic Acid", _
    Array( _
    "7. Acetic Acid_Sl_8", _
    "2. Acetic acid_Sl_18", _
    "2. Acetic acid_Sl_20", _
    "8. Acetic acid_Sl_29", _
    "2. Acetic acid_Sl_38", _
    "2. Acetic acid_Sl_40", _
    "10. Acetic acid_Sl_52", _
    "2. Acetic acid_Sl_61", _
    "2. Acetic acid_Sl_63", _
    "2. Acetic acid_Sl_69", _
    "6. Acetic acid_Sl_75", _
    "2. Acetic acid_Sl_77", _
    "2. Acetic acid_Sl_97", _
    "2. Acetic acid_Sl_103" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Dispersing Agent", _
    Array( _
    "9. Dispersing Agent_Sl_30", _
    "11. Dispersing Agent/Leveling Agent_Sl_53" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Modified Starch", _
    Array( _
    "1. Starch /Modified Starch_Sl_11", _
    "1. Starch /Modified Starch_Sl_31", _
    "1. Starch /Modified Starch_Sl_54", _
    "1. Starch /Modified Starch_Sl_64", _
    "1. Starch /Modified Starch_Sl_86", _
    "1. Starch /Modified Starch_Sl_98" _
    ))

    ' Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    ' "Detergent", _
    ' Array( _

    ' ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Sequestering Agent", _
    Array( _
    "3. Sequestering agent_Sl_3", _
    "3. Sequestering agent_Sl_23", _
    "3. Sequestering agent_Sl_43", _
    "3. Sequestering agent_Sl_72", _
    "2. Sequestering agent_Sl_91" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Fixing Agent", _
    Array( _
    "4.  Fixing Agent_Sl_4", _
    "4. Fixing Agent_Sl_24", _
    "4. Fixing Agent_Sl_44" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Binder", _
    Array( _
    "2. Binder_Sl_12", _
    "2. Binder_Sl_32", _
    "2. Binder_Sl_55", _
    "2. Binder_Sl_65", _
    "2. Binder_Sl_87", _
    "2. Binder_Sl_99" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "PVA", _
    Array( _
    "4. PVA_Sl_14", _
    "4. PVA_Sl_34", _
    "4. PVA_Sl_57", _
    "4. PVA_Sl_67", _
    "4. PVA_Sl_89", _
    "4. PVA_Sl_101" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Wax", _
    Array( _
    "3. Wax_Sl_13", _
    "3. Wax_Sl_33", _
    "3. Wax_Sl_56", _
    "3. Wax_Sl_66", _
    "3. Wax_Sl_88", _
    "3. Wax_Sl_100" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Desizing Agent / Enzyme", _
    Array( _
    "1. Desizing agent_Sl_15", _
    "1. Desizing agent_Sl_35", _
    "1. Desizing agent_Sl_58" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Caustic Soda", _
    Array( _
    "1. Caustic soda (NaOH) Solid_Sl_1", _
    "1. Caustic soda (Solid)_Sl_17", _
    "1. Caustic soda (NaOH)_Sl_21", _
    "1. Caustic soda (Solid)_Sl_37", _
    "1. Caustic soda (NaOH)_Sl_41", _
    "1. Caustic soda (Solid)_Sl_60", _
    "1. Caustic soda (Solid)_Sl_68", _
    "1. Caustic soda (Liquid) , (Concentration active substance not more than 60%)_Sl_70", _
    "3. Caustic Soda (Solid)_Sl_92" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Sodium Hydro Sulphate", _
    Array( _
    "6. Sodium Hydro Sulphite_Sl_27", _
    "9. Sodium Hydro Sulphite_Sl_51" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Sulphuric Acid", _
    Array( _
    "1.98% Sulfuric Acid (Commercial)_Sl_104" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Hydrogen Peroxide", _
    Array( _
    "9. Hydrogent Peroxide_Sl_10", _
    "1. Hydrogent Peroxide (liquid), Concentration 45%_Sl_90" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Stabilizing Agent(Estabilizador FE)", _
    Array( _
    "4. Peroxide Stabilizing Agent_Sl_94" _
    ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Water Decoloring Agent.", _
    Array( _
    "2. Water decoloring Agent_Sl_105" _
    ))

    ' Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    ' "Pumice Stone", _
    ' Array( _

    ' ))

    Set qtyAsGroupDic = Application.Run("dedo_consumption.sumRawMaterialsAsGroupAndAddToDic", qtyAsGroupDic, allDedoConDicAfterAppliedUsedPercentageSpecificRawMaterials, _
    "Stretch Wrapping Film", _
    Array( _
    "Strech Wrapping Flim_Sl_110" _
    ))

    Set finalRawMaterialsQtyCalculatedAsGroup = qtyAsGroupDic

End Function














