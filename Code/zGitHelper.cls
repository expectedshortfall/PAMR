VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zGitHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CODE_PATH As String = "Code\"
Private thisClassModuleName As String

'=============================================================
'
'   PUBLIC
'
'=============================================================
Public Sub ExportModules()
    Dim export As Boolean
    Dim fileName As String
    Dim exportPath As String
    Dim curCommponent As VBIDE.VBComponent
    
    If Not Me.RepoPathExists(exportPath) Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill exportPath & "\*.*"
    On Error GoTo 0
    
    For Each curCommponent In ThisWorkbook.VBProject.VBComponents
        If curCommponent.Type = vbext_ct_Document Or curCommponent.Name = thisClassModuleName Then
            'Thisworkbook or worksheet module or this Module
            'We do nothing
        Else
            export = True
            fileName = curCommponent.Name
    
            Select Case curCommponent.Type
                Case vbext_ct_ClassModule
                    fileName = fileName & ".cls"
                Case vbext_ct_MSForm
                    fileName = fileName & ".frm"
                Case vbext_ct_StdModule
                    fileName = fileName & ".bas"
                Case vbext_ct_Document
                    export = False
            End Select
            
            If export Then
                curCommponent.export exportPath & fileName
            End If
    
            End If
    Next curCommponent
    
    Debug.Print "Export is ready"
    
End Sub

Public Sub ImportModules()
    
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim importPath As String
    Dim curCommponents As VBIDE.VBComponents
    
    If Not Me.RepoPathExists(importPath) Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If
    
    If ThisWorkbook.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
        Exit Sub
    End If
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(importPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If
    
    Call SetupForLoading

    Set curCommponents = ThisWorkbook.VBProject.VBComponents
    
    For Each objFile In objFSO.GetFolder(importPath).Files
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            curCommponents.Import objFile.Path
        End If
    Next objFile
    
    Debug.Print "Import is ready"

End Sub

Function RepoPathExists(ByRef exportPath As String) As Boolean
    Dim WshShell As Object
    Dim FSO As Object
    Dim rootPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    rootPath = ThisWorkbook.Path

    If Right(rootPath, 1) <> "\" Then
        rootPath = rootPath & "\"
    End If
    
    exportPath = rootPath & CODE_PATH
    
    If FSO.FolderExists(exportPath) = False Then
        On Error Resume Next
        MkDir exportPath
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(exportPath) = True Then
        RepoPathExists = True
    Else
        RepoPathExists = False
    End If
            
End Function

'=============================================================
'
'   STRUCTRUAL
'
'=============================================================
Private Sub Class_Initialize()
    'thisClassModuleName = Application.VBE.ActiveCodePane.CodeModule.Name
    thisClassModuleName = "zGitHelper"
End Sub

'=============================================================
'
'   PRIVATE
'
'=============================================================
Private Function SetupForLoading()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Or VBComp.Name = thisClassModuleName Then
                'Thisworkbook or worksheet module or this Module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function
