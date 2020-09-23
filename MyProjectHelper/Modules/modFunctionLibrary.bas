Attribute VB_Name = "modFunctionLibrary"
Option Explicit

Public Sub CentreForm(Optional obj)
   If IsMissing(obj) Then Exit Sub
   
   obj.Move (Screen.Width - obj.Width) / 2, (Screen.Height - obj.Height) / 2
   
End Sub

Public Sub NotDone()
   Dim MsgStr As String
   
   MsgStr = ""
   
   MsgStr = "This function has not yet been implemented."
   
   MsgBox MsgStr, vbInformation
   
End Sub

Public Sub NoHelp()
   Dim MsgStr As String
   
   MsgStr = ""
   
   MsgStr = "There is no Help function associated with this topic."
   
   MsgBox MsgStr, vbInformation

End Sub

Public Sub ShowHourglass()
   Screen.MousePointer = vbHourglass
End Sub

Public Sub ArrowHourglass()
   Screen.MousePointer = vbArrowHourglass
End Sub

Public Sub DefaultPointer()
   Screen.MousePointer = vbDefault
End Sub

Public Sub ResetMousePointer()
   Screen.MousePointer = vbDefault
End Sub

'---------------------------------------------------
'<Purpose> selects all text in a TextBox
'<WhereUsed> within _GotFocus event
'---------------------------------------------------
Public Sub SelectAll(ByVal c As Control)
  c.SelStart = 0
  c.SelLength = Len(c.Text)
End Sub

Public Sub UnselectAll(ByVal c As Control)
   On Error Resume Next
   c.SelStart = 0
   c.SelLength = 0
'   c.Text = UCase(c.Text)
End Sub

Sub SelectEx(ctl As Control, Optional Alignment As AlignmentConstants = vbLeftJustify)
      
   ctl.SelStart = 0
   ctl.SelLength = Len(ctl.Text)
   
   '// Bypass errors
   On Error Resume Next
   
   ctl.Alignment = Alignment
   
   SelectAll ctl
   
End Sub


Sub UnselectEx(ctl As Control, Optional Alignment As AlignmentConstants = vbLeftJustify)
   
   ctl.SelStart = 0
   ctl.SelLength = 0
   
   '// Bypass errors
   On Error Resume Next
   
   ctl.Alignment = Alignment
   
   UnselectAll ctl
   
End Sub

Public Function WriteConfigFile(bstrPathname, bstrIconFilename)
   Dim WrkStr As String
   
   Dim varFilename As String
   
   WrkStr = "": varFilename = ""
   
   WrkStr = "[.ShellClassInfo]" + vbCrLf
   WrkStr = WrkStr + "IconFile=" + bstrIconFilename + vbCrLf
   WrkStr = WrkStr + "IconIndex=0" + vbCrLf
      
   varFilename = bstrPathname
   varFilename = varFilename + "\Desktop.ini"
      
   If FileExists(varFilename) Then
      SetAttr varFilename, vbNormal
      
      NewDoEvents
      
      DeleteTextFile varFilename
      
      NewDoEvents
   End If
   
   Save2File WrkStr, varFilename
   
   NewDoEvents
   
   SetAttr varFilename, vbHidden + vbSystem
   
   NewDoEvents
   
End Function

Public Sub OpenFolderEx(hWndParent As Long, Optional bstrPathname As String = vbNullString)
         
   '0&
   
   If Len(bstrPathname) = 0 Then Exit Sub
   
   Call ShellExecute(hWndParent, _
        "Open", _
        bstrPathname, _
        vbNullString, _
        vbNullString, _
        SW_NORMAL)

End Sub


Public Sub OpenProjectFolderEx(Optional bstrPathname As String = vbNullString)
   Const hWndParent = 0&
   
   If Len(bstrPathname) = 0 Then Exit Sub
   
   Call ShellExecute(hWndParent, _
        "Open", _
        bstrPathname, _
        vbNullString, _
        vbNullString, _
        SW_NORMAL)

End Sub



