Attribute VB_Name = "modGlobal"
Option Explicit

Sub Main()
   Load frmMain
   With frmMain
      .Caption = App.ProductName
      .Show
   End With
   
End Sub
