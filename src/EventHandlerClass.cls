VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventHandlerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents MyApp As Word.Application
Attribute MyApp.VB_VarHelpID = -1

Private Sub MyApp_Quit()
    Main.DeleteFromMenu
End Sub
