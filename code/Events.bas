Public myCharacter As Character

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
    With ThisWorkbook.Sheets("GAME")
        .Range("AY3").value = myCharacter.HP
        .Range("AZ3").value = myCharacter.MP
        .Range("BA3").value = myCharacter.Attack
        .Range("BB3").value = myCharacter.Defense
        .Range("BC3").value = myCharacter.Gold
        .Range("BD3").value = myCharacter.Exp
        .Range("BE3").value = myCharacter.Level
    End With
End Sub