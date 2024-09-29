' Sys Module
Public VBAEnabled As Boolean

Public previousCell As Range
Public previousCellValue As String
Public facingDirection As String

Public charText As String
Public charColor As Long
Public wallText As String
Public wallColor As Long
Public goldText As String
Public goldColor As Long
Public chestText As String
Public chestColor As Long
Public enemyText As String
Public enemyColor As Long
Public specialTexts() As Variant

' Initialize global variables
Sub InitializeGlobals()
    previousCellValue = ""
    facingDirection = "down"

    charText = "@" ' Text representing the character
    charColor = RGB(128, 0, 128) ' Purple color
    wallText = "##" ' Text to prevent selection
    wallColor = RGB(169, 169, 169) ' Dark grey color
    goldText = "$" ' Text representing gold
    goldColor = RGB(255, 215, 0) ' Gold color
    chestText = "[]" ' Text representing a chest
    chestColor = RGB(0, 0, 255) ' Blue color
    enemyText = "E" ' Text representing an enemy
    enemyColor = RGB(255, 0, 0) ' Red color

    ' Store all special texts in an array
    specialTexts = Array(wallText, goldText, chestText, enemyText)
End Sub

' Function to check if a value is in the special texts array
Public Function IsSpecialText(value As String) As Boolean
    Dim i As Integer
    For i = LBound(specialTexts) To UBound(specialTexts)
        If value = specialTexts(i) Then
            IsSpecialText = True
            Exit Function
        End If
    Next i
    IsSpecialText = False
End Function

Public Function PreventSpecificTextSelection(ByVal Target As Range) As Boolean
    ' Check if the selected cell contains any special text
    If Target.value = wallText Or _
       Target.value = chestText Or _
       Target.value = enemyText Then
        PreventSpecificTextSelection = True
    ElseIf Target.value = goldText Then
        ' Add gold to the character's inventory
        myCharacter.Gold = myCharacter.Gold + 1
        ' Update the character stats
        UpdateCharacterStats
        PreventSpecificTextSelection = False
    Else
        PreventSpecificTextSelection = False
    End If
End Function

Public Sub UpdateSelectedCellAddress(ByVal Target As Range)
    ' Display the address of the newly selected cell in cell A1
    ThisWorkbook.Sheets("GAME").Range("A1").value = "Selected Cell: " & Target.Address
    ' Display the facing direction of the character in cell A2
    ThisWorkbook.Sheets("GAME").Range("A2").value = "Facing Direction: " & facingDirection

    ' Log the character's movement
    LogMessage "Moved " & facingDirection & " to " & Target.Address
End Sub

Public Sub MoveChar(ByVal Target As Range)
    ' Prevent selection of cells with specific text
    If PreventSpecificTextSelection(Target) Then
        ' Revert to the previous selection
        Application.EnableEvents = False
        If Not previousCell Is Nothing Then
            previousCell.Select
        End If
        Application.EnableEvents = True
    Else
        SelectCellWithoutScrolling Target

        ' Check if the previous cell is not Nothing
        If Not previousCell Is Nothing Then
            ' Check if the previous cell does not contain special text
            previousCell.value = previousCellValue
        End If

        ' Store the previous cell's value
        previousCell = Target
        previousCellValue = Target.value

        ' Update the character's position
        Target.value = charText

        UpdateSelectedCellAddress Target

        ' Update the previous cell
        Set previousCell = Target
    End If
End Sub

Sub SelectCellWithoutScrolling(targetCell As Range)
    Dim currentScrollRow As Long
    Dim currentScrollColumn As Long
    
    ' Store the current scroll position
    currentScrollRow = ActiveWindow.ScrollRow
    currentScrollColumn = ActiveWindow.ScrollColumn
    
    ' Disable screen updating
    Application.ScreenUpdating = False
    
    ' Select the target cell
    targetCell.Select
    
    ' Restore the previous scroll position
    ActiveWindow.ScrollRow = currentScrollRow
    ActiveWindow.ScrollColumn = currentScrollColumn
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
End Sub

Sub LogMessage(message As String)
    Dim newLogText As String
    Dim logLines() As String
    Dim i As Integer

    ' Reference the cell where the log is stored
    Set logRange = ThisWorkbook.Sheets("GAME").Range("N1")

    ' Read the existing log messages
    logText = logRange.Value

    ' Ensure logText is a string
    If IsEmpty(logText) Then
        logText = ""
    Else
        logText = CStr(logText)
    End If

    ' Prepend the new message to the existing logText
    newLogText = "> " & message & vbCrLf & logText

    ' Split the concatenated logText into an array of lines
    logLines = Split(newLogText, vbCrLf)

    ' Keep only the most recent 10 lines
    If UBound(logLines) > 9 Then
        ReDim Preserve logLines(9)
    End If

    ' Join the array back into a single string
    logText = Join(logLines, vbCrLf)

    ' Write the updated log back to the cell
    logRange.Value = logText
End Sub

Public Sub ClearLog()
    ' Reference the cell where the log is stored
    Set logRange = ThisWorkbook.Sheets("GAME").Range("N1")
    
    ' Clear the log message
    logRange.Value = ""
End Sub