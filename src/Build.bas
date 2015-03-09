Attribute VB_Name = "Build"
Option Explicit

Function GetSourceFiles() As Collection
    Dim out As New Collection
    out.Add "StatusForm.frm"
    out.Add "SettingsForm.frm"
    out.Add "Main.bas"
    out.Add "Compat.bas"
    out.Add "EventHandlerClass.cls"
    out.Add "Build.bas"
    out.Add "Strings.bas"
    Set GetSourceFiles = out
End Function

Function GetLibFiles() As Collection
    Dim out As New Collection
    out.Add "WebResponse.cls"
    out.Add "WebRequest.cls"
    out.Add "WebHelpers.bas"
    out.Add "WebClient.cls"
    out.Add "IWebAuthenticator.cls"
    out.Add "Dictionary.cls"
    Set GetLibFiles = out
End Function

Function GetDocumentDir() As String
    Dim dir As String
    dir = ActiveDocument.path
    If Right(dir, 5) = ".dotm" Then
        dir = Left(dir, InStrRev(dir, Application.PathSeparator) - 1)
    End If
    GetDocumentDir = dir
End Function

Function FileNameWithoutExtensionFromPath(ByRef strFullPath As String) As String
    If InStrRev(strFullPath, ".") > 0 Then
         FileNameWithoutExtensionFromPath = Left(strFullPath, InStrRev(strFullPath, ".") - 1)
    Else
        FileNameWithoutExtensionFromPath = strFullPath
    End If
End Function

Function GetProject(d As Document) As VBProject
    #If Mac Then
    #Else
        Set GetProject = d.VBProject
        Exit Function
    #End If
    
    Dim p As VBProject
    Dim foo As VBProjects
    Set foo = Application.VBE.VBProjects
    For Each p In Application.VBE.VBProjects
    Debug.Print (FileNameWithoutExtensionFromPath(p.fileName) & " -- " & FileNameWithoutExtensionFromPath(d.FullName) & " -- " & FileNameWithoutExtensionFromPath(p.BuildFileName))
        If FileNameWithoutExtensionFromPath(p.fileName) = FileNameWithoutExtensionFromPath(d.FullName) Or _
                FileNameWithoutExtensionFromPath(p.BuildFileName) = FileNameWithoutExtensionFromPath(d.FullName) Then
            Set GetProject = p
            Exit Function
        End If
    Next
End Function

Function GetComponentByFileName(project As VBProject, fileName As String) As VBComponent
    Dim Name As String
    Dim VBComp As VBComponent
    Name = FileNameWithoutExtensionFromPath(fileName)
    
    For Each VBComp In project.VBComponents
        If VBComp.Name = Name Then
            Set GetComponentByFileName = VBComp
            Exit Function
        End If
    Next VBComp
End Function

Sub ImportModules()
    If MsgBox(Prompt:="Really overwrite modules with code files on disk?", Buttons:=vbOKCancel) <> vbOK Then
        Exit Sub
    End If

    Dim project As VBProject
    Set project = GetProject(Application.ActiveDocument)
    
    If project Is Nothing Then
        MsgBox "Error: Can't identify VBProject. Try saving, closing, and re-opening this file, and make sure it is the only document open."
        Exit Sub
    End If
    
    Dim fileName As Variant
    For Each fileName In GetSourceFiles()
        ImportModule project, GetDocumentDir() & Application.PathSeparator & "src", CStr(fileName)
    Next fileName
    For Each fileName In GetLibFiles()
        ImportModule project, GetDocumentDir() & Application.PathSeparator & "lib", CStr(fileName)
    Next fileName
    
    ' TODO: Can't seem to rename module from itself.
    Dim buildComponent As VBComponent
    Set buildComponent = GetComponentByFileName(project, "Build1")
    buildComponent.Name = "Build"
    
    MsgBox "Done. Please rename the 'Build1' Module to 'Build'."
End Sub

Function ImportModule(project As VBProject, path As String, fileName As String)
    Dim VBComp As VBComponent
    Set VBComp = GetComponentByFileName(project, fileName)
    If Not VBComp Is Nothing Then
        project.VBComponents.Remove VBComp
    End If
    
    project.VBComponents.Import path & Application.PathSeparator & fileName
End Function

Sub ExportModules()
    If MsgBox(Prompt:="Really overwrite code files on disk with project modules?", Buttons:=vbOKCancel) <> vbOK Then
        Exit Sub
    End If

    Dim project As VBProject
    Set project = GetProject(Application.ActiveDocument)
    
    If project Is Nothing Then
        MsgBox "Error: Can't identify VBProject. Try saving, closing, and re-opening this file, and make sure it is the only document open."
        Exit Sub
    End If
    
    Dim fileName As Variant
    For Each fileName In GetSourceFiles()
        ExportModule project, GetDocumentDir() & Application.PathSeparator & "src", CStr(fileName)
    Next fileName
    For Each fileName In GetLibFiles()
        ExportModule project, GetDocumentDir() & Application.PathSeparator & "lib", CStr(fileName)
    Next fileName
    
    MsgBox "Done."
End Sub

Function ExportModule(project As VBProject, path As String, fileName As String)
    Dim VBComp As VBComponent
    Set VBComp = GetComponentByFileName(project, fileName)
    VBComp.Export path & Application.PathSeparator & fileName
End Function


