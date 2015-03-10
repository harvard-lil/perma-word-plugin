Attribute VB_Name = "Compat"
Option Explicit

#If Mac Then
#ElseIf VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
    Private Declare PtrSafe Function apiCopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
    Private Declare Function apiCopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
#End If


Public Function CompatCopyFile(source As String, dest As String)
    ' copy-file function capable of copying open files
    #If Mac Then
        'MacScript ("tell application ""Finder"" to copy file """ & source & """ to folder """ & dest & """")
        Dim result As ShellResult
        result = WebHelpers.ExecuteInShell("cp " & WebHelpers.PrepareTextForShell(GetPOSIXPath(source)) & " " & WebHelpers.PrepareTextForShell(GetPOSIXPath(dest)))
        If result.ExitCode <> 0 Then
            err.Raise result.ExitCode, "CopyFile", result.Output
        End If
        
    #Else
        Dim ExitCode As Long
        ExitCode = apiCopyFile(source, dest, False)
        If ExitCode = 0 Then
            err.Raise 0, "CopyFile", "Copy failed."
        End If
    #End If
End Function

Function CompatSleep(seconds As Double)
    ' sleep function that doesn't use CPU cycles
    #If Mac Then
        WebHelpers.ExecuteInShell ("sleep " & CStr(seconds))
    #Else
        Sleep CLng(seconds * 1000)
    #End If
End Function

Function CompatFileExists(ByVal fileName As String) As Boolean
    On Error GoTo Catch
    FileSystem.FileLen fileName
    CompatFileExists = True
    GoTo Finally
Catch:
        CompatFileExists = False
Finally:
End Function

Function GetPOSIXPath(MacPath As String) As String
    #If Mac Then
        GetPOSIXPath = MacScript("return (POSIX path of """ & MacPath & """) as String")
    #Else
        GetPOSIXPath = "[undefined]"
    #End If
End Function

Function SetSystemFont(form As UserForm)
    Dim SystemFont As String
    Dim c As Control
    
    #If Mac Then
        SystemFont = "Lucida Grande"
    #Else
        SystemFont = "Sugoe UI"
    #End If
    
    For Each c In form.Controls
        c.Font.Name = SystemFont
    Next c
End Function
