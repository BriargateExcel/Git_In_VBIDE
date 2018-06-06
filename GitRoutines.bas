Attribute VB_Name = "GitRoutines"
Option Explicit

Private EventHandlers As New Collection

Public Sub CreateMenu()
    Dim MenuEvent As GitHandlerClass
    Dim HelpMenu As CommandBarControl
    Dim NewMenu As CommandBarPopup
    Dim MenuItem As CommandBarControl

    '   Delete the menu and event handlers if they already exist
    DeleteMenuAndEventHandlers
    
    '   Find the Help Menu
    Set HelpMenu = Application.VBE.CommandBars(1).FindControl(ID:=30010)
    
    If HelpMenu Is Nothing Then
        '       Add the menu to the end
        Set NewMenu = Application.VBE.CommandBars(1).Controls.Add _
                      (Type:=msoControlPopup, _
                       temporary:=True)
    Else
        '      Add the menu before Help
        Set NewMenu = Application.VBE.CommandBars(1).Controls.Add _
                      (Type:=msoControlPopup, _
                       Before:=HelpMenu.Index, _
                       temporary:=True)
    End If

    '   Add a caption for the menu
    NewMenu.Caption = "&GitCommands"
    
    '   ADD NEW MENU ITEM
    Set MenuEvent = New GitHandlerClass
    Set MenuItem = NewMenu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "Export VBA Project"
        .OnAction = "'" & ThisWorkbook.Name & "'!ExportVBAProject"
    End With
    Set MenuEvent.GitEvtHandler = Application.VBE.Events.CommandBarEvents(MenuItem)
    EventHandlers.Add MenuEvent
    
    '   ADD NEW MENU ITEM
    Set MenuEvent = New GitHandlerClass
    Set MenuItem = NewMenu.Controls.Add(Type:=msoControlButton)
    With MenuItem
        .Caption = "Import VBA Project"
        .OnAction = "ImportVBAProject"
    End With
    Set MenuEvent.GitEvtHandler = Application.VBE.Events.CommandBarEvents(MenuItem)
    EventHandlers.Add MenuEvent
    
End Sub

'#################################################
Private Sub DeleteMenuAndEventHandlers()
    Dim I As Long
    On Error Resume Next
    Application.VBE.CommandBars(1).Controls("GitCommands").Delete
    For I = 0 To EventHandlers.Count - 1
        EventHandlers.Remove 1
    Next I
    On Error GoTo 0
End Sub

Public Sub ExportVBAProject()
    Dim Export As Boolean
    Dim FileName As String
    Dim cmpComponent As VBIDE.VBComponent
    Dim Path As String
    Dim VBAProj As Variant
    Dim ProjectToExport As VBIDE.VBProject

    GitForm.ProjectList.Clear
    
    For Each VBAProj In Application.VBE.VBProjects
        GitForm.ProjectList.AddItem VBAProj.Name
    Next VBAProj
    
    GitForm.ProjectList.Text = GitForm.ProjectList.List(0)
    
    GitForm.Show
    
    Set ProjectToExport = Application.VBE.VBProjects(GitForm.ProjectList.Value)

    If ProjectToExport.Protection = 1 Then
        MsgBox "This project is protected, not possible to export the code"
        Exit Sub
    End If

    Path = GetFolder( _
           "Select the Folder to Place the VBA Code", _
           ThisWorkbook.Path)
    If Path = vbNullString Then Exit Sub
    Path = Path & "\"

    For Each cmpComponent In ProjectToExport.VBComponents

        Export = True
        FileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
        Case vbext_ct_ClassModule
            FileName = FileName & ".cls"
        Case vbext_ct_MSForm
            FileName = FileName & ".frm"
        Case vbext_ct_StdModule
            FileName = FileName & ".bas"
        Case vbext_ct_Document
            '                FileName = FileName & ".cls"
            Export = False
        Case vbext_ct_ActiveXDesigner
            Export = False
        End Select

        If Export Then
            ''' Export the component to a text file.
            cmpComponent.Export Path & FileName

        End If

    Next cmpComponent

    MsgBox "All Forms, Modules, and Classes have been exported into the " & _
           Path & " folder"

End Sub                                          ' ExportVBAProject

Private Function GetFolder( _
        ByVal Descrip As String, _
        ByVal StartingPath As String _
        ) As String
    ' VBNullString return means user cancelled
    
    If StartingPath <> "|" Then StartingPath = StartingPath & "\"
    
    Dim Folder As FileDialog
    Dim SelItem As String
    Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
    With Folder
        .Title = Descrip
        .AllowMultiSelect = False
        .InitialFileName = StartingPath
        If .Show <> -1 Then GoTo NextCode
        SelItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = SelItem
    Set Folder = Nothing
End Function

Public Sub ImportVBAProject()

    GitForm.ProjectList.Clear

    Dim VBAProj As Variant
    For Each VBAProj In Application.VBE.VBProjects
        GitForm.ProjectList.AddItem VBAProj.Name
    Next VBAProj
    
    GitForm.ProjectList.Text = GitForm.ProjectList.List(0)

    GitForm.Show

    Dim ProjectToImport As VBIDE.VBProject
    Set ProjectToImport = Application.VBE.VBProjects(GitForm.ProjectList.Value)
    Unload GitForm
    If DeleteComponentsInProject(ProjectToImport) Then
        ' components deleted; proceed
    Else
        ' components not deleted; stop processing
        Exit Sub
    End If
    
    If ProjectToImport.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected; " & _
               "it is not possible to export the code", _
               vbOKOnly Or vbCritical, _
               "Protected workbook"
        Exit Sub
    End If

    'Get the path to the folder with modules
    Dim Path As String
    Path = GetFolder( _
           "Select the Folder Containing the VBA Code You Want to Import", _
           ThisWorkbook.Path)
    If Path = vbNullString Then Exit Sub
    If FolderWithVBAProjectFiles(Path) = "Error" Then
        MsgBox "Import Folder does not exist", _
               vbOKOnly Or vbCritical, _
               "No Import Folder"
        Exit Sub
    End If

    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    If FSO.GetFolder(Path).Files.Count = 0 Then
        MsgBox "There are no files to import", _
               vbOKOnly Or vbCritical, _
               "No Files"
        Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    DeleteComponentsInProject ProjectToImport

    Dim cmpComponents As VBIDE.VBComponents
    Set cmpComponents = ProjectToImport.VBComponents
    Dim ScriptFile As Scripting.File

    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each ScriptFile In FSO.GetFolder(Path).Files
    
        If (FSO.GetExtensionName(ScriptFile.Name) = "cls") Or _
            (FSO.GetExtensionName(ScriptFile.Name) = "frm") Or _
             (FSO.GetExtensionName(ScriptFile.Name) = "bas") Then
            cmpComponents.Import ScriptFile.Path
        End If
    
    Next ScriptFile

    MsgBox "All Forms, Modules, and Classes have been imported", _
           vbOKOnly Or vbInformation, _
           "Import Complete"

End Sub                                          ' ImportVBAProject

Private Function FolderWithVBAProjectFiles(ByRef Path As String) As String
    '
    '*******************************************************************************
    ' Function Name:FolderWithVBAProjectFiles
    '
    ' Function Purpose:
    ' Returns a string containing the full path name of the folder where the Forms,
    ' Modules, and Classes will be copied to or from.
    ' The folder is called VBAProjectFiles.
    ' It creates the folder if it doesn't exist.
    '
    ' Called by:
    ' ExportModules
    ' ImportModules
    '
    ' Return value:
    ' String containing the full path name of the VBAProjectFiles folder
    '
    '*******************************************************************************
    '
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    If Not FSO.FolderExists(Path) Then
        On Error Resume Next
        MkDir Path
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(Path) Then
        FolderWithVBAProjectFiles = Path
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function                                     ' FolderWithVBAProjectFiles

Private Function DeleteComponentsInProject(ByVal Proj As VBIDE.VBProject) As Boolean
    ' True = components deleted
    ' False = components not deleted

    DeleteComponentsInProject = False
    Select Case MsgBox("All components of this VBA Project will be deleted. " & _
                       "Do you want to continue?", _
                       vbYesNo Or vbQuestion, _
                       "Delete Components")

    Case vbYes
        Dim VBComp As VBIDE.VBComponent
    
        For Each VBComp In Proj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'This is a workbook or worksheet module, we do nothing
            Else
                Proj.VBComponents.Remove VBComp
            End If
        Next VBComp
        DeleteComponentsInProject = True
    Case vbNo
        DeleteComponentsInProject = False
    End Select

End Function                                     ' DeleteComponentsInProject


