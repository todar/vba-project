VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''
' This VBProject exports all VBComponent source code
' after each save event. This can be changed per project,
' but would be a highly recommended workflow.
'
' @author Robert Todar <robert@roberttodar.com>
' @ref {Module} Project
''
Option Explicit

' Export all code after every save.
' Ideally in a project you would only want to do this with developers.
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Project.ExportComponentsToSourceFolders
    Project.ExportReferencesToConfigFile
End Sub
