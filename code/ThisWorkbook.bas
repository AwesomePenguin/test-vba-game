Private Sub Workbook_Open()
    VBAEnabled = True

    ' Clear the log message when the workbook is opened
    Call ClearLog

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

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Clear the log message before the workbook is closed
    Call ClearLog
End Sub
