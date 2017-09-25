Attribute VB_Name = "DevTools"
'''''''''''''''''''''''''''''''''''''''''''''''''''
''''Desarrollado:       Juan Antonio Barragán  ''''
''''Email:              jabarragann@unal.edu.co''''
'''''''''''''''''''''''''''''''''''''''''''''''''''

Sub import()

    Dim path As String
    path = Application.ActiveWorkbook.path & "\SourceCode\"
    
    importSourceFiles path

End Sub

Sub export()

    Dim path As String
    path = Application.ActiveWorkbook.path & "\SourceCode\"
    
    removeFilesInSourceDirectory
    exportSourceFiles path

End Sub
Public Sub removeFilesInSourceDirectory()

    Dim objFSO As Object, objFolder As Object, folder As Object
    Dim path As String
    
    path = Application.ActiveWorkbook.path
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(path)
    
    If objFSO.FolderExists(path & "/SourceCode/") Then
        objFSO.DeleteFolder (path & "/SourceCode"), True
         MkDir path & "\SourceCode"
    End If
    
   
    
    
End Sub
Public Sub exportSourceFiles(destPath As String)
 
    Dim component As VBComponent
    Dim objFSO As Object, objFolder As Object, folder As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If Not objFSO.FolderExists(destPath) Then
         MkDir destPath
    End If
    
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            component.export destPath & component.Name & ToFileExtension(component.Type)
        End If
    Next
 
End Sub

Public Sub importSourceFiles(sourcePath As String)
    Dim file As String
    file = Dir(sourcePath)
    
    While (file <> vbNullString)
        
        If file <> "DevTools.bas" Then
            Application.VBE.ActiveVBProject.VBComponents.import sourcePath & file
        End If
        
        file = Dir
    Wend
    
End Sub

 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
    Case vbext_ComponentType.vbext_ct_ClassModule
    ToFileExtension = ".cls"
    Case vbext_ComponentType.vbext_ct_StdModule
    ToFileExtension = ".bas"
    Case vbext_ComponentType.vbext_ct_MSForm
    ToFileExtension = ".frm"
    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
    Case vbext_ComponentType.vbext_ct_Document
    Case Else
    ToFileExtension = vbNullString
    End Select
 
End Function

Public Sub removeAllModules()
    Dim project As VBProject
    Set project = Application.VBE.ActiveVBProject
     
    Dim comp As VBComponent
    
    For Each comp In project.VBComponents
        If Not comp.Name = "DevTools" And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
            project.VBComponents.Remove comp
        End If
    Next
    
End Sub




