Attribute VB_Name = "utility_formating_fun"
Option Explicit

Private Function rangeFormat(appliedRange As Range, fontName As String, fontSize As Integer, isFontBold As Boolean, isWrapText As Boolean, hAlignment As Variant, vAlignment As Variant, numFormat As String)

    With appliedRange
        .Font.Name = fontName
        .Font.Size = fontSize
        .Font.Bold = isFontBold
        .WrapText = isWrapText
        .HorizontalAlignment = hAlignment
        .VerticalAlignment = vAlignment
        .NumberFormat = numFormat
    End With

End Function

Private Function SetBorderInsideHairlineAroundThin(appliedRange As Range)

        Application.Run "utility_formating_fun.setBorder", appliedRange, xlEdgeTop, xlThin
        Application.Run "utility_formating_fun.setBorder", appliedRange, xlEdgeRight, xlThin
        Application.Run "utility_formating_fun.setBorder", appliedRange, xlEdgeBottom, xlThin
        Application.Run "utility_formating_fun.setBorder", appliedRange, xlEdgeLeft, xlThin
        Application.Run "utility_formating_fun.setBorder", appliedRange, xlInsideHorizontal, xlHairline
        Application.Run "utility_formating_fun.setBorder", appliedRange, xlInsideVertical, xlHairline

    
    
End Function

Private Function setBorder(rng As Range, appliedSide As Variant, borderWeight As Variant)

    With rng.Borders(appliedSide)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = borderWeight
    End With

End Function

Private Function removeBorder(rng As Range, appliedSide As Variant)

    rng.Borders(appliedSide).LineStyle = xlNone
        
End Function






















''may be next time no need below function
'Private Function borderAroundEachCellsThin(appliedRange As Range)
'
'    Dim appliedCells As Range
'
'    For Each appliedCells In appliedRange
'        appliedCells.BorderAround LineStyle:=xlContinuous, Weight:=xlThin
'    Next appliedCells
'
'End Function
'
'Private Function borderAroundEachCellsHairline(appliedRange As Range)
'
'    Dim appliedCells As Range
'
'    For Each appliedCells In appliedRange
'        appliedCells.BorderAround LineStyle:=xlContinuous, Weight:=xlHairline
'    Next appliedCells
'
'End Function
'
'Private Function borderAroundEachCellsNone(appliedRange As Range)
'
'    Dim appliedCells As Range
'
'    For Each appliedCells In appliedRange
'        appliedCells.Borders(xlDiagonalDown).LineStyle = xlNone
'        appliedCells.Borders(xlDiagonalUp).LineStyle = xlNone
'        appliedCells.Borders(xlEdgeLeft).LineStyle = xlNone
'        appliedCells.Borders(xlEdgeTop).LineStyle = xlNone
'        appliedCells.Borders(xlEdgeBottom).LineStyle = xlNone
'        appliedCells.Borders(xlEdgeRight).LineStyle = xlNone
'        appliedCells.Borders(xlInsideVertical).LineStyle = xlNone
'        appliedCells.Borders(xlInsideHorizontal).LineStyle = xlNone
'
'    Next appliedCells
'
'End Function
'
'
'Private Function borderAroundThin(appliedRange As Range)
'
'    appliedRange.BorderAround LineStyle:=xlContinuous, Weight:=xlThin
'
'End Function
'
'Private Function borderAroundHairline(appliedRange As Range)
'
'    appliedRange.BorderAround LineStyle:=xlContinuous, Weight:=xlHairline
'
'End Function
'
'Private Function borderEdgeTopHairline(appliedRange As Range)
'
'    With appliedRange.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'    End With
'
'End Function
'
'Private Function borderEdgeRightHairline(appliedRange As Range)
'
'    With appliedRange.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'    End With
'
'End Function
'
'Private Function borderEdgeBottomHairline(appliedRange As Range)
'
'    With appliedRange.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'    End With
'
'End Function
'
'Private Function borderEdgeLeftHairline(appliedRange As Range)
'
'    With appliedRange.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'    End With
'
'End Function
'
'Private Function borderEdgeTopThin(appliedRange As Range)
'
'    With appliedRange.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'    End With
'
'End Function
'
'Private Function borderEdgeRightThin(appliedRange As Range)
'
'    With appliedRange.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'    End With
'
'End Function
'
'Private Function borderEdgeBottomThin(appliedRange As Range)
'
'    With appliedRange.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'    End With
'
'End Function
'
'Private Function borderEdgeLeftThin(appliedRange As Range)
'
'    With appliedRange.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'    End With
'
'End Function

