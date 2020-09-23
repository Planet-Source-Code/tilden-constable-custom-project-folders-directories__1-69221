Attribute VB_Name = "modFreeFiles"
Option Explicit

'Author: PulseWave
'Author's email: PulseWave@aol.com
'Date Submitted: 2/24/1999
'Compatibility: VB 6,VB 5,VB 4/32

'Task: Save and Load text files

'Declarations

'Code:
Public Function GetFileData(ByVal bstrFilename As String) As String
   '// Declare variables
   Dim nResult, a
   Dim WrkStr As String
   
   '// Initialize
   WrkStr = ""
   
   On Error GoTo ErrHandler
   
   Open bstrFilename For Input As #1
   
   Do While Not EOF(1)
      Line Input #1, a
      WrkStr = WrkStr + a + Chr(13) + Chr(10)
   Loop
   
   'data = WrkStr
   
   Close #1
   
   GetFileData = WrkStr
   
   Exit Function

ErrHandler:
   If Err.Number <> 0 Then
      '// Re-initialize
      WrkStr = ""
      
      '// Re-initialize
      WrkStr = "An error has occured!" & vbCrLf & vbCrLf
      WrkStr = WrkStr & Err.Number & ": " & Err.Description
      
      '// Display error message.
      MsgBox WrkStr, vbExclamation
      
      '// Clear errors
      Err.Clear
   
      GetFileData = ""
   
   End If
   
End Function

Public Sub Save2File(Optional data As String = vbNullString, Optional bstrFilename As String = vbNullString)
   '// Declare variables
   Dim nResult, a
   Dim WrkStr As String
   
   If Len(bstrFilename) = 0 Then Exit Sub
   
   '// Initialize
   WrkStr = ""
   
   On Error GoTo ErrHandler
   
   Open bstrFilename For Output As #1
      Print #1, data
   Close 1
   
   Exit Sub

ErrHandler:
   If Err.Number <> 0 Then
      '// Re-initialize
      WrkStr = ""
      
      '// Re-initialize
      WrkStr = "An error has occured!" & vbCrLf & vbCrLf
      WrkStr = WrkStr & Err.Number & ": " & Err.Description
      
      '// Display error message.
      MsgBox WrkStr, vbExclamation
      
      '// Clear errors
      Err.Clear
   
   End If

End Sub


