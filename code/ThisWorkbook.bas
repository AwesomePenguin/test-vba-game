Private Sub Workbook_Open()
    VBAEnabled = True

    Application.DisplayFullScreen = True
    ThisWorkbook.Sheets("GAME").Activate
    ActiveWindow.Zoom = 90

    ' Initialize global variables
    Sys.InitializeGlobals
    
    ' Initialize the previous cell
    Set previousCell = ActiveCell

    Events.InitializeCharacter
    
    ' Assign arrow keys to movement functions
    Application.OnKey "{UP}", "HandleUpKey"
    Application.OnKey "{LEFT}", "HandleLeftKey"
    Application.OnKey "{DOWN}", "HandleDownKey"
    Application.OnKey "{RIGHT}", "HandleRightKey"
End Sub

