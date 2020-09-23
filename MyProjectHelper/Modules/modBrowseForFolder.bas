Attribute VB_Name = "modBrowseForFolder"
Option Explicit

Private Const BIF_STATUSTEXT = &H4&
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_NEWFOLDERBUTTON = (WM_USER + 100)
Private Const BFFM_SETSELECTIONA = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW = (WM_USER + 103)
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const lPtr = (LMEM_FIXED Or LMEM_ZEROINIT)
Private Const BFFM_INITIALIZED = 1

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpsTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long                  'added

Public mhWnd As Long


Public mvarCurrentDirectory As String

Public Function BrowseForFolder(ByVal StartDir As String, hWnd As Long, Title As String) As String
   Dim lpIDList As Long
   Dim szTitle As String
   Dim sBuffer As String
   Dim tBrowseInfo As BrowseInfo
   Dim lpPath
   
   mvarCurrentDirectory = StartDir '& vbNullChar
  
   szTitle = Title
   
   With tBrowseInfo
      .hWndOwner = hWnd
      .lpsTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BFFM_NEWFOLDERBUTTON
      .lpfnCallback = FARPROC(AddressOf BrowseCallbackProcStr)
      
      lpPath = LocalAlloc(lPtr, Len(StartDir) + 1)
'
'      CopyMemory ByVal lpPath, ByVal StartDir, Len(StartDir) + 1
'
      .lParam = lpPath

   End With
   
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      BrowseForFolder = sBuffer
   Else
      BrowseForFolder = ""
   End If
   
End Function

Public Function OpenFolder(sPath As String, lHwnd As Long, Optional sTitle As String = "Selecione a pasta desejada") As String
Dim lpIDList As Long
Dim sBuffer As String
Dim lpPath As Long
Dim tBrowseInfo As BrowseInfo

With tBrowseInfo
   .hWndOwner = lHwnd
   .pIDLRoot = 0
   .lpsTitle = lstrcat(sTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BFFM_NEWFOLDERBUTTON
   .lpfnCallback = FARPROC(AddressOf BrowseCallbackProcStr)
    lpPath = LocalAlloc(lPtr, Len(sPath) + 1)
    CopyMemory ByVal lpPath, ByVal sPath, Len(sPath) + 1
    .lParam = lpPath
End With

lpIDList = SHBrowseForFolder(tBrowseInfo)

If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End If

OpenFolder = sBuffer
End Function


Private Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
   Select Case uMsg
      Case BFFM_INITIALIZED
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
   End Select
End Function

Private Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function

Private Function HiWord(ByVal dwValue As Long) As Long
  Dim hexstr As String
    hexstr = Right$("00000000" & Hex$(dwValue), 8)
    HiWord = CLng("&H" & Left$(hexstr, 4))
End Function

Private Function LoWord(ByVal dwValue As Long) As Long
  Dim hexstr As String
    hexstr = Right$("00000000" & Hex$(dwValue), 8)
    LoWord = CLng("&H" & Right$(hexstr, 4))
End Function

Private Sub SwapByte(byte1 As Byte, byte2 As Byte)
    byte1 = byte1 Xor byte2
    byte2 = byte1 Xor byte2
    byte1 = byte1 Xor byte2
End Sub

Private Function FixedHex(ByVal hexval As Long, ByVal nDigits As Long) As String
    FixedHex = Right$("00000000" & Hex$(hexval), nDigits)
End Function


'Private Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
'   Dim lpIDList As Long
'   Dim ret As Long
'   Dim sBuffer As String
'
'   On Error Resume Next
'
'   mhWnd = hWnd
'
'   Select Case uMsg
'      Case BFFM_INITIALIZED
'         Call SendMessage(hWnd, BFFM_SETSELECTION, 1, mvarCurrentDirectory)
'
'      Case BFFM_SELCHANGED
'         sBuffer = Space(MAX_PATH)
'         ret = SHGetPathFromIDList(lp, sBuffer)
'
'         If ret = 1 Then
'            Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
'         End If
'
'   End Select
'
'   BrowseCallbackProc = 0
'
'End Function
'
'Private Function GetAddressOfFunction(Add As Long) As Long
'   GetAddressOfFunction = Add
'End Function
'
'
'Private Function FARPROC(pfn As Long) As Long
'  FARPROC = pfn
'End Function
'
'Private Function HiWord(ByVal dwValue As Long) As Long
'  Dim hexstr As String
'    hexstr = right$("00000000" & Hex$(dwValue), 8)
'    HiWord = CLng("&H" & left$(hexstr, 4))
'End Function
'
'Private Function LoWord(ByVal dwValue As Long) As Long
'  Dim hexstr As String
'    hexstr = right$("00000000" & Hex$(dwValue), 8)
'    LoWord = CLng("&H" & right$(hexstr, 4))
'End Function

'Private Sub SwapByte(byte1 As Byte, byte2 As Byte)
'    byte1 = byte1 Xor byte2
'    byte2 = byte1 Xor byte2
'    byte1 = byte1 Xor byte2
'End Sub
'
'Private Function FixedHex(ByVal hexval As Long, ByVal nDigits As Long) As String
'    FixedHex = right$("00000000" & Hex$(hexval), nDigits)
'End Function

Public Function GetFileName(FileName As String) As String
    Dim I As Integer
    Dim tmp As String
    For I = 1 To Len(FileName)
        tmp = Right$(FileName, I)
        If Left(tmp, 1) = "\" Then
            GetFileName = Mid(tmp, 2)
            Exit Function
        End If
    Next
   GetFileName = ""
End Function

Public Function GetFileExtension(FileName As String, Optional LowerCase As Boolean = True) As String
    If (LowerCase) Then
        GetFileExtension = LCase(Right$(FileName, 4))
    Else
        GetFileExtension = Right$(FileName, 4)
    End If
End Function

Public Function GetFilePath(FileName As String, Optional IncludeDrive As Boolean = True) As String
    Dim I As Integer
    Dim str As String
    For I = 1 To Len(FileName)
        str = Right$(FileName, I)
        If Mid(str, 1, 1) = "\" Then
            Dim iLenght As Integer
            If (IncludeDrive) Then iLenght = 1 Else iLenght = 4
            GetFilePath = Mid(FileName, iLenght, Len(FileName) - I) & "\"
            Exit Function
        End If
    Next
    GetFilePath = ""
End Function

Public Function GetDrive(FileName As String, Optional IncludeSlash As Boolean = False) As String
    Dim iLenght As Integer
    If (IncludeSlash) Then iLenght = 3 Else iLenght = 2
    GetDrive = LCase(Left$(FileName, iLenght))
End Function

Public Function FixPath(sPath As String) As String
   FixPath = sPath & IIf(Right(sPath, 1) <> "\", "\", "")
End Function


'*************************************************
' Open an Explorer browser on the specified directory path.
' Simply supply the Me.hWnd from your form for the hWndParent value.
'*************************************************
Public Function BrowsePath(hWndParent As Long, Path As String) As Long
  BrowsePath = ShellPath(hWndParent, "explore", Path)
End Function

'*************************************************
' support routine for module
'*************************************************
Private Function ShellPath(hWnd As Long, Cmd As String, Path As String) As Long
  Dim I As Long
  I = ShellExecute(hWnd, Cmd, Path, "", "", SW_NORMAL)
  If I = 0 Then I = 1
  If I > 32 Then I = 0
  ShellPath = I
End Function



