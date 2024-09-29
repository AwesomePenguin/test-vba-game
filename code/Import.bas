Sub ImportVBAFromDirectory()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim FilePath As String
    Dim FileName As String
    Dim CompName As String
    Dim ThisModuleName As String
    Dim FileNum As Integer
    Dim FileContent As String
    
    ' Set the directory path where your VBA files are stored
    FilePath = "C:\Users\Si Hang\OneDrive\Documents\Coding Projects\VBA\test-vba-game\code\"
    
    ' Get the VBA project of the workbook
    Set vbProj = ThisWorkbook.VBProject
    
    ' Get the name of this module to prevent importing itself
    ThisModuleName = "Import" ' Change this to the actual name of your import module if different
    
    ' Import all .bas files (modules)
    FileName = Dir(FilePath & "*.bas")
    Do While FileName <> ""
        CompName = Left(FileName, InStrRev(FileName, ".") - 1)
        ' Skip importing this module itself
        If CompName <> ThisModuleName Then
            ' Check if the file is not empty
            FileNum = FreeFile
            Open FilePath & FileName For Input As FileNum
            FileContent = Input(LOF(FileNum), FileNum)
            Close FileNum
            If Len(FileContent) > 0 Then
                ' Check if the component already exists and remove it
                On Error Resume Next
                Set vbComp = vbProj.VBComponents(CompName)
                If Not vbComp Is Nothing Then
                    vbProj.VBComponents.Remove vbComp
                End If
                On Error GoTo 0
                ' Import the new component
                Set vbComp = vbProj.VBComponents.Import(FilePath & FileName)
                ' Rename the component to match the file name
                vbComp.Name = CompName
            End If
        End If
        FileName = Dir
    Loop
    
    ' Import all .cls files (classes)
    FileName = Dir(FilePath & "*.cls")
    Do While FileName <> ""
        CompName = Left(FileName, InStrRev(FileName, ".") - 1)
        ' Skip importing this module itself
        If CompName <> ThisModuleName Then
            ' Check if the file is not empty
            FileNum = FreeFile
            Open FilePath & FileName For Input As FileNum
            FileContent = Input(LOF(FileNum), FileNum)
            Close FileNum
            If Len(FileContent) > 0 Then
                ' Check if the component already exists and remove it
                On Error Resume Next
                Set vbComp = vbProj.VBComponents(CompName)
                If Not vbComp Is Nothing Then
                    vbProj.VBComponents.Remove vbComp
                End If
                On Error GoTo 0
                ' Import the new component
                Set vbComp = vbProj.VBComponents.Import(FilePath & FileName)
                ' Rename the component to match the file name
                vbComp.Name = CompName
            End If
        End If
        FileName = Dir
    Loop
    
    ' Import all .frm files (forms)
    FileName = Dir(FilePath & "*.frm")
    Do While FileName <> ""
        CompName = Left(FileName, InStrRev(FileName, ".") - 1)
        ' Skip importing this module itself
        If CompName <> ThisModuleName Then
            ' Check if the file is not empty
            FileNum = FreeFile
            Open FilePath & FileName For Input As FileNum
            FileContent = Input(LOF(FileNum), FileNum)
            Close FileNum
            If Len(FileContent) > 0 Then
                ' Check if the component already exists and remove it
                On Error Resume Next
                Set vbComp = vbProj.VBComponents(CompName)
                If Not vbComp Is Nothing Then
                    vbProj.VBComponents.Remove vbComp
                End If
                On Error GoTo 0
                ' Import the new component
                Set vbComp = vbProj.VBComponents.Import(FilePath & FileName)
                ' Rename the component to match the file name
                vbComp.Name = CompName
                ' Check for associated .frx file and copy it
                If Dir(FilePath & CompName & ".frx") <> "" Then
                    FileCopy FilePath & CompName & ".frx", ThisWorkbook.Path & "\" & CompName & ".frx"
                End If
            End If
        End If
        FileName = Dir
    Loop
    
    MsgBox "VBA code imported successfully!"
End Sub