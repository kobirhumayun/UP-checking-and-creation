Attribute VB_Name = "JsonUtilityFunction"
'Sub saveFile()
'
'    Dim dict As Object
'    Set dict = CreateObject("Scripting.Dictionary")
'    ' Populate the dictionary with some data
'    Dim i As Long
'    For i = 0 To 10
'        dict.Add "Name" & i, "Test Name" & i
'        dict.Add "Age" & i, i
'        dict.Add "City" & i, "Dhaka"
'    Next i
'
'    ' Specify the file path
'    SaveDictionaryToJsonTextFile dict, ActiveWorkbook.path & Application.PathSeparator & "file.json"
'
'
'End Sub
'
'Sub loadFile()
'
'    Dim dict As Object
'
'    Set dict = LoadDictionaryFromJsonTextFile(ActiveWorkbook.path & Application.PathSeparator & "file.json")
'
'    ' Access dictionary items
'    Dim key As Variant
'    For Each key In dict.Keys
'        Debug.Print key
'    Next key
'
'End Sub

Private Function SaveDictionaryToJsonTextFile(dict As Object, filePath As String)

    ' Convert the dictionary to JSON
    Dim json As String
    json = JsonConverter.ConvertToJson(dict)
    
    ' Write the JSON to a file
    Open filePath For Output As #1
    Print #1, json
    Close #1

    Debug.Print "Dictionary save as Json"

End Function

Private Function LoadDictionaryFromJsonTextFile(filePath As String) As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim json As String

    ' Read the JSON from the file
    Open filePath For Input As #1
    json = Input(LOF(1), #1)
    Close #1
    
    ' Convert JSON to dictionary
    Set dict = JsonConverter.ParseJson(json)
    
    Set LoadDictionaryFromJsonTextFile = dict

End Function


