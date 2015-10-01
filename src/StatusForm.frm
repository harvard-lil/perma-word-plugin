VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatusForm 
   Caption         =   "Creating Perma Archive"
   ClientHeight    =   3000
   ClientLeft      =   -560
   ClientTop       =   -11500
   ClientWidth     =   7060
   OleObjectBlob   =   "StatusForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StatusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Activate()
    Me.Repaint
    Main.InsertPermaLink
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    Compat.SetSystemFont Me
End Sub
