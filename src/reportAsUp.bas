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
