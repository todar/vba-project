Attribute VB_Name = "Project"
''
' !! THIS MODULE IS NOT IMPORTED, PLEASE ONLY MODIFY IN THE VB EDITOR !!
'
' This module is created to easily export and import VBA code
' into and from `./src` and `./tests` directories.
'
' This is needed as Git can't read Excel files directly, but can read the
' source files that are exported.
'
' The export should be ran either on `Workbook_AfterSave` event. Example:
'
'   Private Sub Workbook_AfterSave(ByVal Success As Boolean)
'       ExportComponentsToSourceFolders
'   End Sub
'
'
' Note, this module can't use functions or code from other modules
' as this module will delete those for importing and won't be able to be accessed.
'
'
' @author Robert Todar <robert@roberttodar.com>
' @status Development - Exporting seems to be working fine right now,
'                       but still having issues with importing. Looks Like
'                       I can import everything but Sheets, ThisWorkbook, and .frx files.
'                       Maybe create a solution to store just the text of code for sheets
'                       and workbook vs the code export.
' @ref {Microsoft Visual Basic For Applications Extensibility 5.3} VBComponets
' @ref {Microsoft Visual Basic For Applications Extensibility 5.3} VBComponet
' @ref {Microsoft Scripting Runtime} Scripting.FileSystemObject
''
Option Explicit
Option Compare Text

Private Const PATTERN_FOR_TEST_MODULES As String = "*_Tests"
Private Const REFERENCES_FILE_NAME As String = "references.txt"

' Root Directory of this Project.
Public Property Get Dirname() As String
    Dirname = ThisWorkbook.path
End Property

' Directory where all source code will be stored. `./src`
Public Property Get SourceDirectory() As String
    SourceDirectory = joinPaths(Dirname, "src")
End Property

' Directory where all tests for the source code will be stored. `./tests`
Public Property Get TestsDirectory() As String
    TestsDirectory = joinPaths(Dirname, "tests")
End Property

' This Projects VB thisProjectsVBComponents.
' @NOTE: Should this be a single project, or should I use this
'        for any project/workbook? For now will leave as the
'        current
Private Property Get thisProjectsVBComponents() As VBComponents
    Set thisProjectsVBComponents = ThisWorkbook.VBProject.VBComponents
End Property

' Get the file extension for a VBComponent. That is the component name and the proper extension.
Private Function getVBComponentFilename(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule
            getVBComponentFilename = component.Name & ".cls"
            
        Case vbext_ComponentType.vbext_ct_StdModule
            getVBComponentFilename = component.Name & ".bas"
            
        Case vbext_ComponentType.vbext_ct_MSForm
            getVBComponentFilename = component.Name & ".frm"
            
        Case vbext_ComponentType.vbext_ct_Document
            getVBComponentFilename = component.Name & ".cls"
            
        Case Else
            ' @TODO: Need to think of possible throwing an error?
            ' Is it possible to get something else?? I don't think so
            ' Will need to double check this.
            Debug.Print "Unknown component"
    End Select
End Function

' Check to see if component exits in this current Project
Private Function componentExists(ByVal filename As String) As Boolean
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)
        
        If getVBComponentFilename(component) = filename Then
            componentExists = True
            Exit Function
        End If
    Next index
End Function

' Export all modules in this current workbook into a src dir
Public Sub ExportComponentsToSourceFolders()
    ' Make sure the source directory exists before adding to it.
    createCleanDirectory SourceDirectory
    createCleanDirectory TestsDirectory
    
    ' Loop each component within this project and export the correct directory.
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)
        
        ' Export component to the correct source directory
        ' using the name of the component and the correct file extension.
        ' Soure files will either go into the main source folder, or the
        ' tests folder depending on the naming convention.
        component.Export joinPaths( _
                                    IIf(component.Name Like PATTERN_FOR_TEST_MODULES, TestsDirectory, SourceDirectory), _
                                    getVBComponentFilename(component) _
                                  )
    Next index
End Sub

Private Sub createCleanDirectory(folderpath As String)
    ' Create folder if it doesn't already exist.
    Dim fso As New Scripting.FileSystemObject
    If Not fso.FolderExists(folderpath) Then
        fso.CreateFolder folderpath
    End If
    
    ' Delete any files within the directory to clean it out.
    Dim file As file
    For Each file In fso.GetFolder(folderpath).Files
        file.Delete
    Next file
End Sub

' Import source code from the source Directory.
' This works by first deleting all current components,
' then importing all the components from the source directory.
'
' @status Testing && Development
' @warn This will cause files to overwrite that already exists.
' @warn This will also remove files not found in the source component.
Public Sub DangerouslyImportComponentsFromSourceFolders()
    Dim fso As New Scripting.FileSystemObject
    
    ' Remove current components to make room for the imported ones.
    Dim file As file
    For Each file In fso.GetFolder(SourceDirectory).Files
        ' If the component already, it needs to be deleted in order to
        ' import the file, otherwise an error is thrown.
        If componentExists(file.Name) And file.Name <> "Project.bas" Then
            Dim component As VBComponent
            Set component = thisProjectsVBComponents.Item(fso.GetBaseName(file.Name))
            
            ' Unable to remove document type components (Sheets, workbook)
            If component.Type <> vbext_ct_Document Then
                ' This removes the component but doesn't from memory until
                ' after all code execution has completed.
                thisProjectsVBComponents.Remove component
            End If
        End If
    Next file
    
    ' After all code is finished executing, the components removed above will
    ' finally be removed from memory.
    Application.OnTime Now, "saftleyImportAfterCleanup"
End Sub

Private Sub saftleyImportAfterCleanup()
    Dim fso As New Scripting.FileSystemObject

    Dim file As file
    For Each file In fso.GetFolder(SourceDirectory).Files
        If Not componentExists(file.Name) And fso.GetExtensionName(file.Name) <> "frx" Then
            ' Safe to import the source file as there are no conflicts of names.
            thisProjectsVBComponents.Import joinPaths(SourceDirectory, file.Name)
        End If
    Next file
    
    ' Adding for tests as well
    For Each file In fso.GetFolder(TestsDirectory).Files
        If Not componentExists(file.Name) And fso.GetExtensionName(file.Name) <> "frx" Then
            ' Safe to import the source file as there are no conflicts of names.
            thisProjectsVBComponents.Import joinPaths(TestsDirectory, file.Name)
        End If
    Next file
End Sub

' Converts the VBComponent enum to a string representation of type of component.
Private Function getVBComponentTypeName(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule
            getVBComponentTypeName = "Class Module"
            
        Case vbext_ComponentType.vbext_ct_StdModule
            getVBComponentTypeName = "Module"
            
        Case vbext_ComponentType.vbext_ct_MSForm
            getVBComponentTypeName = "Form"
            
        Case vbext_ComponentType.vbext_ct_Document
            getVBComponentTypeName = "Document"
            
        Case Else
            ' All components should be accounted for, this is just in case ;)
            Debug.Print "Unknown type: " & component.Type
    End Select
End Function

' Prints out details about all VBComponents in the current project
' @status Development
Private Sub printDiffFromSourceFolder()
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)
        
        Debug.Print getVBComponentFilename(component)
    Next index
End Sub

' Helper function to join paths...
Private Function joinPaths(ParamArray paths() As Variant) As String
    Dim fso As New Scripting.FileSystemObject
    Dim index As Long
    For index = LBound(paths) To UBound(paths)
        joinPaths = fso.BuildPath(joinPaths, Replace(paths(index), "/", "\"))
    Next
End Function

' Export All Project Library References into a text file for reimporting.
Public Sub ExportReferencesToConfigFile()
    Dim myProject As VBProject
    Set myProject = Application.VBE.ActiveVBProject
    
    Dim fso As New Scripting.FileSystemObject
    
    With fso.OpenTextFile(joinPaths(Dirname, REFERENCES_FILE_NAME), ForWriting, True)
        Dim library As Reference
        For Each library In myProject.References
            .WriteLine library.Name & vbTab & library.GUID & vbTab & library.Major & vbTab & library.Minor
        Next
    End With
End Sub

' Import from Config file to update all needed references.
Public Sub ImportReferencesFromConfigFile()
    Dim fso As New Scripting.FileSystemObject
    
    With fso.OpenTextFile(joinPaths(Dirname, REFERENCES_FILE_NAME), ForReading, True)
        Dim line As Long
        Do While Not .AtEndOfStream
            Dim values As Variant
            values = Split(.ReadLine, vbTab)
            
            ' Just skip if it already exists
            On Error Resume Next
            ThisWorkbook.VBProject.References.AddFromGuid values(1), values(2), values(3)
        Loop
    End With
End Sub
