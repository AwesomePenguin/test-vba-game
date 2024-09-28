' Test.bas
Sub TestCharacter()
    ' Ensure the character is initialized
    If myCharacter Is Nothing Then
        InitializeCharacter
    End If
    
    ' Display initial attributes in the Immediate Window
    Debug.Print "HP: " & myCharacter.HP
    Debug.Print "MP: " & myCharacter.MP
    Debug.Print "Attack: " & myCharacter.Attack
    Debug.Print "Defense: " & myCharacter.Defense
    
    ' Display initial attributes in the worksheet
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
    End With
    
    ' Character takes damage
    myCharacter.TakeDamage 30
    Debug.Print "HP after taking damage: " & myCharacter.HP
    
    ' Update HP in the worksheet
    ThisWorkbook.Sheets("MAP").Range("H3").Value = myCharacter.HP
    
    ' Character uses magic
    myCharacter.UseMagic 20
    Debug.Print "MP after using magic: " & myCharacter.MP
    
    ' Update MP in the worksheet
    ThisWorkbook.Sheets("MAP").Range("I3").Value = myCharacter.MP
End Sub