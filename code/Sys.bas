' Sys Module
Public VBAEnabled As Boolean

Public previousCell As Range
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

Public myCharacter As Character

' Initialize global variables
Sub InitializeGlobals()
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
    If Target.Value = wallText Or _
       Target.Value = chestText Or _
       Target.Value = enemyText Then
        PreventSpecificTextSelection = True
    ElseIf Target.Value = goldText Then
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
    ThisWorkbook.Sheets("MAP").Range("A1").Value = "Selected Cell: " & Target.Address
End Sub

Public Sub HighlightSelectedCell(ByVal Target As Range)
    ' Check if the previous cell is not Nothing
    If Not previousCell Is Nothing Then
        ' Check if the previous cell does not contain special text
        If Not IsSpecialText(previousCell.Value) Then
            ' Clear the highlight from the previously selected cell
            previousCell.Interior.ColorIndex = xlNone
        Else
            Select Case previousCell.Value
                Case wallText
                    previousCell.Interior.Color = wallColor
                Case goldText
                    previousCell.Interior.Color = goldColor
                Case chestText
                    previousCell.Interior.Color = chestColor
                Case enemyText
                    previousCell.Interior.Color = enemyColor
            End Select
        End If
    End If
    
    ' Highlight the newly selected cell as purple
    Target.Interior.Color = charColor
End Sub

Public Sub HighlightSpecialTextCells()
    Dim ws As Worksheet
    Dim cell As Range
    
    ' Target the "MAP" sheet
    Set ws = ThisWorkbook.Sheets("MAP")
    
    ' Loop through each cell in the used range of the "MAP" sheet
    For Each cell In ws.UsedRange
        ' Check if the cell contains special text
        If IsSpecialText(cell.Value) Then
            ' Apply the corresponding color
            Select Case cell.Value
                Case wallText
                    cell.Interior.Color = wallColor
                Case goldText
                    cell.Interior.Color = goldColor ' Yellow color
                Case chestText
                    cell.Interior.Color = chestColor ' Blue color
                Case enemyText
                    cell.Interior.Color = enemyColor ' Red color
            End Select
        End If
    Next cell
End Sub

Public Sub InitializeCharacter()
    Set myCharacter = New Character
    
    ' Set initial attributes
    myCharacter.HP = 100
    myCharacter.MP = 50
    myCharacter.Attack = 20
    myCharacter.Defense = 10
    myCharacter.Gold = 0
End Sub

Public Sub UpdateCharacterStats()
    ' Ensure the character is initialized
    If myCharacter Is Nothing Then
        InitializeCharacter
    End If
    
    ' Display character stats in the worksheet
    With ThisWorkbook.Sheets("MAP")
        .Range("H1").Value = "Character Stats"
        .Range("H2").Value = "HP"
        .Range("H3").Value = myCharacter.HP
        .Range("I2").Value = "MP"
        .Range("I3").Value = myCharacter.MP
        .Range("J2").Value = "ATK"
        .Range("J3").Value = myCharacter.Attack
        .Range("K2").Value = "DEF"
        .Range("K3").Value = myCharacter.Defense
        .Range("L2").Value = "Gold"
        .Range("L3").Value = myCharacter.Gold
        .Range("M2").Value = "Exp"
        .Range("M3").Value = myCharacter.Exp
        .Range("N2").Value = "Level"
        .Range("N3").Value = myCharacter.Level
    End With
End Sub