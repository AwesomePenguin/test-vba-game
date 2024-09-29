' HandleKeys.bas
' This module is responsible for handling key presses in the game.
Private isHandlingKey As Boolean

Public Sub HandleUpKey()
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    facingDirection = "up"
    Dim targetCell As Range
    Set targetCell = ActiveCell.Offset(-1, 0)
    Sys.MoveChar targetCell
    isHandlingKey = False
End Sub

Public Sub HandleDownKey()
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    facingDirection = "down"
    Dim targetCell As Range
    Set targetCell = ActiveCell.Offset(1, 0)
    Sys.MoveChar targetCell
    isHandlingKey = False
End Sub

Public Sub HandleLeftKey()
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    facingDirection = "left"
    Dim targetCell As Range
    Set targetCell = ActiveCell.Offset(0, -1)
    Sys.MoveChar targetCell
    isHandlingKey = False
End Sub

Public Sub HandleRightKey()
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    facingDirection = "right"
    Dim targetCell As Range
    Set targetCell = ActiveCell.Offset(0, 1)
    Sys.MoveChar targetCell
    isHandlingKey = False
End Sub