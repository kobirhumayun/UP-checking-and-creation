Attribute VB_Name = "readUp"
Option Explicit

Private Function readUpAsDict(upWs As Worksheet) As Object

    Dim upAsDict As Object
    Set upAsDict = CreateObject("Scripting.Dictionary")
        
    upAsDict.Add "upClause1", Application.Run("readUp.upClause1AsDict", upWs)
    upAsDict.Add "upClause6", Application.Run("readUp.upClause6AsDict", upWs)
    upAsDict.Add "upClause7", Application.Run("readUp.upClause7AsDict", upWs)
    upAsDict.Add "upClause8", Application.Run("readUp.upClause8AsDict", upWs)
    upAsDict.Add "upClause9", Application.Run("readUp.upClause9AsDict", upWs)
    upAsDict.Add "upClause11", Application.Run("readUp.upClause11AsDict", upWs)
    upAsDict.Add "upClause12a", Application.Run("readUp.upClause12aAsDict", upWs)
    upAsDict.Add "upClause12bFabrics", Application.Run("readUp.upClause12bFabricsAsDict", upWs)
    upAsDict.Add "upClause12bGarments", Application.Run("readUp.upClause12bGarmentsAsDict", upWs)
    upAsDict.Add "upClause13", Application.Run("readUp.upClause13AsDict", upWs)
    upAsDict.Add "upClause14", Application.Run("readUp.upClause14AsDict", upWs)
    
    
    Set readUpAsDict = upAsDict
    
End Function

Private Function upClause1AsDict(upWs As Worksheet) As Object

    Dim clause1AsDict As Object
    Set clause1AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause1AsDict = clause1AsDict
    
End Function

Private Function upClause6AsDict(upWs As Worksheet) As Object

    Dim clause6AsDict As Object
    Set clause6AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause6AsDict = clause6AsDict
    
End Function

Private Function upClause7AsDict(upWs As Worksheet) As Object

    Dim clause7AsDict As Object
    Set clause7AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause7AsDict = clause7AsDict
    
End Function

Private Function upClause8AsDict(upWs As Worksheet) As Object

    Dim clause8AsDict As Object
    Set clause8AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause8AsDict = clause8AsDict
    
End Function

Private Function upClause9AsDict(upWs As Worksheet) As Object

    Dim clause9AsDict As Object
    Set clause9AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause9AsDict = clause9AsDict
    
End Function

Private Function upClause11AsDict(upWs As Worksheet) As Object

    Dim clause11AsDict As Object
    Set clause11AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause11AsDict = clause11AsDict
    
End Function

Private Function upClause12aAsDict(upWs As Worksheet) As Object

    Dim clause12aAsDict As Object
    Set clause12aAsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause12aAsDict = clause12aAsDict
    
End Function

Private Function upClause12bFabricsAsDict(upWs As Worksheet) As Object

    Dim clause12bFabricsAsDict As Object
    Set clause12bFabricsAsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause12bFabricsAsDict = clause12bFabricsAsDict
    
End Function

Private Function upClause12bGarmentsAsDict(upWs As Worksheet) As Object

    Dim clause12bGarmentsAsDict As Object
    Set clause12bGarmentsAsDict = CreateObject("Scripting.Dictionary")
    
    Set upClause12bGarmentsAsDict = clause12bGarmentsAsDict
    
End Function

Private Function upClause13AsDict(upWs As Worksheet) As Object

    Dim clause13AsDict As Object
    Set clause13AsDict = CreateObject("Scripting.Dictionary")
        
    Set upClause13AsDict = clause13AsDict
    
End Function

Private Function upClause14AsDict(upWs As Worksheet) As Object

    Dim clause14AsDict As Object
    Set clause14AsDict = CreateObject("Scripting.Dictionary")
    
    Set upClause14AsDict = clause14AsDict
    
End Function

