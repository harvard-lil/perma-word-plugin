VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "Perma Settings"
   ClientHeight    =   5500
   ClientLeft      =   -2320
   ClientTop       =   -20240
   ClientWidth     =   6900
   OleObjectBlob   =   "SettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub Cancel_Click()
    Unload Me
End Sub


Private Sub Save_Click()
    On Error GoTo HandleError
    
    With Me
        MySaveSetting "APIKey", .APIKey.Value
    End With
    Unload Me
    
Exit Sub
HandleError:
    ShowError err, "Save_Click"
End Sub


Private Sub UninstallButton_Click()
    On Error GoTo HandleError
    
    If MsgBox(Prompt:="Really uninstall the Perma plugin?", Buttons:=vbOKCancel) <> vbOK Then
        Exit Sub
    End If
    
    Unload Me
    DeleteFromMenu
    
    Dim target As addin
    For Each target In Application.AddIns
        If target.Name = "Perma Word Plugin.dotm" Then
            Dim targetFile As String
            targetFile = target.path & Application.PathSeparator & target.Name
            On Error GoTo HandleFileError
            SetAttr targetFile, vbNormal
            Kill targetFile
            MsgBox "Plugin uninstalled."
            Exit Sub
        End If
    Next target
    MsgBox "Error: Unable to find add-in to uninstall. Please quit Word and remove the 'Perma Word Plugin' add-in from your Startup folder manually."

Exit Sub
HandleError:
    ShowError err, "UninstallButton_Click"
Exit Sub
HandleFileError:
    MsgBox "Almost done! To finish uninstalling, remove" & vbCrLf & targetFile & vbCrLf & "and restart Word."
End Sub

Private Sub UserForm_Initialize()
    'On Error GoTo HandleError
    
    With Me
        .APIKey.Value = MyGetSetting("APIKey")
    End With
    
    Compat.SetSystemFont Me
Exit Sub
HandleError:
    ShowError err, "UserForm_Initialize"
End Sub

Private Sub ViewToolsPage_Click()
    On Error GoTo HandleError
    
    ActiveDocument.FollowHyperlink "http://perma.cc/settings/tools/"

Exit Sub
HandleError:
    ShowError err, "ViewToolsPage_Click"
End Sub
