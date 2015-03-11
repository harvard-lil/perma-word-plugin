Attribute VB_Name = "Main"
Option Explicit

Private Const VERSION = "0.0.2"
Private Const InstalledPluginName = "Perma Word Plugin.dotm"

Dim EventHandlers As New EventHandlerClass
Sub AutoExec()
    AddToMenu
    Set EventHandlers.MyApp = Word.Application
End Sub

Sub ShowError(err As ErrObject, context As String)
    MsgBox ("Oops! The Perma plugin encountered an unexpected error. Please let us know the following:" & vbCrLf & vbCrLf & _
                    "- Word Version: " & Application.System.OperatingSystem & " " & Application.VERSION & vbCrLf & _
                    "- Perma Plugin Version: " & VERSION & vbCrLf & _
                    "- Error: " & CStr(err.Number) & " in " & context & ": " & err.Description)
End Sub

Sub InstallPlugin()
    On Error GoTo HandleFileError
    Dim targetFile As String
    targetFile = Application.StartupPath & Application.PathSeparator & InstalledPluginName
    'If Compat.CompatFileExists(targetFile) Then
    '    Kill targetFile
   ' End If
    Debug.Print ("Copying " & ActiveDocument.AttachedTemplate.FullName & " to " & targetFile)
    Compat.CompatCopyFile ActiveDocument.AttachedTemplate.FullName, targetFile
    MsgBox "The Perma plugin is installed! Please restart Word and right-click on a link to get started."
Exit Sub
HandleFileError:
    MsgBox "Sorry, we don't have permission to copy files to your Startup folder. Please copy this plugin to" & vbCrLf & vbCrLf & _
        Application.StartupPath & vbCrLf & vbCrLf & _
        "and restart Word.", vbExclamation
End Sub

Sub DebugShowContextBarNames()
    ' handy for finding the name of other context menus that should have Perma option included
    Dim cmd As CommandBar
    On Error Resume Next
    For Each cmd In Application.CommandBars
        With cmd.Controls.Add(Type:=msoControlButton)
            .OnAction = ""
            .FaceId = 100
            .Caption = cmd.Name
            .BeginGroup = True
            .Tag = "Perma_Tag"
        End With
nextcmd:
    Next cmd
End Sub


Sub AddToMenu()
    On Error GoTo HandleError
    
    ' avoid duplicates.
    DeleteFromMenu
    
    ' DebugShowContextBarNames

    AddSubMenu "Hyperlink Context Menu"
    AddSubMenu "Text"
    AddSubMenu "Spelling"
    AddSubMenu "Grammar"
    AddSubMenu "Table Text"
    AddSubMenu "Text w/ Thesaurus"
    AddSubMenu "Footnotes"

    ' include in Hyperlink submenu
    AddSubMenu "Hyperlink Menu", True
    
Exit Sub
HandleError:
    ShowError err, "AddToMenu"
End Sub

Sub AddSubMenu(BarName As String, Optional InsertOnly As Boolean = False)
    Dim ContextMenu As CommandBar
    
    On Error GoTo BarNotFound
    Set ContextMenu = Application.CommandBars(BarName)
    On Error GoTo HandleError
    
    'With ContextMenu.Controls.Add(Type:=msoControlPopup)
    '    .Caption = "Perma"
    '    .Tag = "Perma_Tag"

        With ContextMenu.Controls.Add(Type:=msoControlButton)
            .OnAction = "ShowInsertPermaLinkForm"
            .FaceId = 1576
            .Caption = "Insert Perma Link..."
            .BeginGroup = True
            .Tag = "Perma_Tag"
        End With
        
        If Not InsertOnly Then
            With ContextMenu.Controls.Add(Type:=msoControlButton)
                .OnAction = "ShowSettings"
                .Caption = "Perma Settings..."
                .Tag = "Perma_Tag"
            End With
        End If
        
    'End With
Exit Sub
BarNotFound:
    ' Debug.Print ("WARNING: Can't find menu bar " & BarName)
Exit Sub
HandleError:
    ShowError err, "AddSubMenu"
End Sub

Sub DeleteFromMenu()
    On Error GoTo HandleError
    
    Dim cmd As CommandBar
    Dim ctrl As CommandBarControl
    
    For Each cmd In Application.CommandBars
        For Each ctrl In cmd.Controls
            If ctrl.Tag = "Perma_Tag" Then
                ctrl.Delete
            End If
        Next ctrl
    Next

Exit Sub
HandleError:
    ShowError err, "DeleteFromMenu"
End Sub


Sub ShowInsertPermaLinkForm()
    StatusForm.Show
End Sub

Sub InsertPermaLink()
    On Error GoTo HandleError
    
    Dim Url As String
    Dim hl As Hyperlink
    
    ' ----- check API key -----
    Dim APIKey As String
    APIKey = MyGetSetting("APIKey")
    If APIKey = "" Then
        If MsgBox(Prompt:="Your Perma API key must be set before you can insert links. Open settings now?", Buttons:=vbOKCancel) = vbOK Then
            SettingsForm.Show
        End If
        Exit Sub
    End If
    
    ' ----- check selected hyperlink -----
    If Selection.Hyperlinks.Count >= 1 Then
        Set hl = Selection.Hyperlinks(1)
        Url = Selection.Hyperlinks(1).Address
    Else
        If Selection.Type = wdSelectionIP Then
           MsgBox Prompt:="Select a hyperlink."
           Exit Sub
        ElseIf Selection.Type <> wdSelectionNormal Then
           MsgBox Prompt:="Not a valid selection."
           Exit Sub
        Else
            Selection.MoveEndWhile Chr(32), wdBackward  ' remove trailing spaces
            Url = Selection.Text
        End If
    End If
    
    ' ----- update status -----
    StatusForm.StatusMessage.Caption = "Creating Perma Link ..."
    StatusForm.StatusBar.Left = StatusForm.StatusBarBackground.Left
    StatusForm.CancelButton.Enabled = False
    StatusForm.Repaint
    DoEvents
    
    ' ----- POST to create endpoint -----
    Dim Client As New WebClient
    Dim Resource As String
    Dim Response As WebResponse
    Client.BaseUrl = "https://dashboard.perma.cc/api/v1/"
    Resource = "archives/?" & _
        "api_key=" & APIKey
    Dim Body As New Dictionary
    Body.Add "url", Url
    Body.Add "title", "foo"
    Set Response = Client.PostJson(Resource, Body)

    ' Debug.Print (ConvertToJson(Response.Data))
    
    ' ----- handle successful link creation -----
    If Response.StatusCode = WebStatusCode.Created Then
    
        Dim newLink As String
        newLink = "http://perma.cc/" & Response.Data("guid")

        StatusForm.StatusMessage.Caption = "Archiving link contents ..."
        StatusForm.CancelButton.Enabled = True

        Dim statusCount, statusOffset, statusChunks As Integer
        statusCount = 0
        statusChunks = 10
        'Dim timeout As Date
        'timeout = DateAdd("s", 120, Now)
        Dim canceled, failed As Boolean
        canceled = False
        failed = False
        
        Do
            ' ----- update status bar -----
            statusCount = statusCount + 1
            statusOffset = statusCount Mod (statusChunks * 2 - 2)
            If statusOffset >= statusChunks Then
                statusOffset = (statusChunks * 2 - 2) - statusOffset
            End If
            StatusForm.StatusBar.Left = StatusForm.StatusBarBackground.Left + (StatusForm.StatusBarBackground.Width / statusChunks) * statusOffset
            StatusForm.Repaint
        
            ' ----- check current status of link -----
            Dim StatusResource As String
            StatusResource = "archives/" & Response.Data("guid") & "/?" & _
                "api_key=" & APIKey
            Set Response = Client.GetJson(StatusResource)
            
            Dim assets As Dictionary
            Set assets = Response.Data("assets")(1)
            If assets("image_capture") <> "pending" Then
                If assets("image_capture") = "failed" And (assets("pdf_capture") = "failed" Or assets("warc_capture") = "failed") Then
                    failed = True
                End If
                
                Exit Do
            End If
            
            ' ----- check for cancellation -----
            DoEvents
            If Not StatusForm.Visible Then
                canceled = True
                Exit Do
            End If
            
            Compat.CompatSleep (1)
            
        Loop While True ' Until Now > timeout
        
        ' ---- write new link into document -----
        SafeStartUndoRecord "Insert Perma Link"
        If Not hl Is Nothing Then
            hl.Range.Select
        End If
        Selection.Collapse wdCollapseEnd
        Selection.Font.Reset  ' close hyperlink if any
        Selection.TypeText Text:=" ["
        Selection.Hyperlinks.Add Anchor:=Selection.Range, _
            Address:=newLink, _
            TextToDisplay:="" & newLink  ' concatenating "" with newLink avoids a weird out-of-memory bug in Mac Word 2011
        If failed Then
            Selection.TypeText Text:=" - FAILED"
        ElseIf canceled Then
            Selection.TypeText Text:=" - NOT CHECKED"
        End If
        Selection.TypeText Text:="]"
        SafeEndUndoRecord
        
    ' handle 401 unauthorized
    ElseIf Response.StatusCode = 401 Then
        MsgBox "Link creation failed. Please check that your API key is correct."
        
    ' handle 408 timeout
    ElseIf Response.StatusCode = 408 Then
        MsgBox "Unable to reach the Perma service."
        
    ' handle 400 validation error
    ElseIf Response.StatusCode = 400 And Not Response.Data Is Nothing Then
        Dim Item
        For Each Item In Response.Data("archives").items
            MsgBox "Error creating Perma Link: " & Item
        Next Item
        
    ' handle 404 error -- API endpoint no longer exists
    ElseIf Response.StatusCode = 404 Then
        MsgBox "This plugin is no longer supported by the Perma service. Please check for a new version of the plugin at https://perma.cc/"
        
    ' catchall error
    Else
        MsgBox "Error " & CStr(Response.StatusCode) & ": " & Response.StatusDescription
    End If
    
Exit Sub
HandleError:
    ShowError err, "InsertPermaLink"
End Sub

Sub ShowSettings()
    SettingsForm.Show
End Sub

Function MyGetSetting(key As String) As String
    On Error GoTo SettingNotFound
    MyGetSetting = GetSetting("Perma", "Settings", key)
    Exit Function
SettingNotFound:
    SaveSetting "Perma", "Settings", key, ""
    MyGetSetting = ""
End Function

Function MySaveSetting(key As String, Value As String)
    SaveSetting "Perma", "Settings", key, Value
End Function

Function SafeStartUndoRecord(RecordName As String)
    #If VBA7 Then
        ' only available in Windows Word 2010+
        On Error Resume Next
        Application.UndoRecord.StartCustomRecord RecordName
    #End If
End Function


Function SafeEndUndoRecord()
    #If VBA7 Then
        ' only available in Windows Word 2010+
        On Error Resume Next
        Application.UndoRecord.EndCustomRecord
    #End If
End Function
    
