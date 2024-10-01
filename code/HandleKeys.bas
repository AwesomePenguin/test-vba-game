Public Sub HandleUpKey()
    If ThisWorkbook.ActiveSheet.Name <> "GAME" Then Exit Sub
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    facingDirection = "up"
    Call Sys.MoveChar(-1, 0)
    isHandlingKey = False
End Sub

Public Sub HandleDownKey()
    If ThisWorkbook.ActiveSheet.Name <> "GAME" Then Exit Sub
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    facingDirection = "down"
    Call Sys.MoveChar(1, 0)
    isHandlingKey = False
End Sub

Public Sub HandleLeftKey()
    If ThisWorkbook.ActiveSheet.Name <> "GAME" Then Exit Sub
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    facingDirection = "left"
    Call Sys.MoveChar(0, -1)
    isHandlingKey = False
End Sub

Public Sub HandleRightKey()
    If ThisWorkbook.ActiveSheet.Name <> "GAME" Then Exit Sub
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    facingDirection = "right"
    Call Sys.MoveChar(0, 1)
    isHandlingKey = False
End Sub

Public Sub HandleMenuKey()
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    ' Switch to the MENU sheet
    ThisWorkbook.Sheets("MENU").Activate
    isHandlingKey = False
End Sub

Public Sub HandleGameKey()
    If isHandlingKey Then Exit Sub
    isHandlingKey = True
    ' Switch to the GAME sheet
    ThisWorkbook.Sheets("GAME").Activate
    isHandlingKey = False
End Sub