Attribute VB_Name = "Strings"
Option Explicit

Function FileExtensionFromPath(ByRef strFullPath As String) As String
     FileExtensionFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "."))
End Function

Function FileNameWithoutExtensionFromPath(ByRef strFullPath As String) As String
     FileNameWithoutExtensionFromPath = Left(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "."))
End Function

Function FolderFromPath(ByRef strFullPath As String) As String
     FolderFromPath = Left(strFullPath, InStrRev(strFullPath, Application.PathSeparator))
End Function

Function FileNameFromPath(strFullPath As String) As String
     FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, Application.PathSeparator))
End Function
