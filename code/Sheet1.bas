Dim hitPoint As Integer
Dim magicPoint As Integer
Dim Attack As Integer
Dim Defense As Integer

Dim Gold As Integer
Dim Exp As Integer

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Check if VBA is enabled
    VBAEnabled = False
    If Not VBAEnabled Then Exit Sub

    ' Prevent selection of cells with specific text
    If Sys.PreventSpecificTextSelection(Target) Then
        ' Revert to the previous selection
        Application.EnableEvents = False
        If Not previousCell Is Nothing Then
            previousCell.Select
        End If
        Application.EnableEvents = True
    Else
        ' Update the current address of the selected cell
        Sys.UpdateSelectedCellAddress Target

        Sys.UpdateCharacterStats
        
        ' Highlight only the selected cell as purple
        Sys.HighlightSelectedCell Target
        
        ' Update the previous cell
        Set previousCell = Target
    End If
End Sub

Private Sub Workbook_Open()
    VBAEnabled = True

    ' Initialize global variables
    Sys.InitializeGlobals
    
    ' Initialize the previous cell
    Set previousCell = ActiveCell
    
    ' Highlight cells containing wall text
    Sys.HighlightSpecialTextCells

    Sys.InitializeCharacter
End Sub