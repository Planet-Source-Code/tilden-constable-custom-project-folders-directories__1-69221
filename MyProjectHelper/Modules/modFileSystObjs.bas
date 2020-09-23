Attribute VB_Name = "modFileSystObjs"
Option Explicit

'*************************************************
' API call used by BrowsePath
'*************************************************
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Declare System Object variables
Public m_objFileSystemObject As Object
Public m_objTextFile As Object
Public m_objFolder As Object
Public m_objDrive As Object
Public m_varFileName As Variant
Public m_varFolderName As Variant
Public m_strMessage As String
Public m_strDateCreated As String
Public m_strFileName As String

Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public FILE_NOT_FOUND As Boolean

Public Const SW_NORMAL = 1
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6



'
' Public Constants
'

Public Const gstrSEP_DIR$ = "\"                         ' Directory separator character
Public Const gstrAT$ = "@"
Public Const gstrSEP_DRIVE$ = ":"                       ' Driver separater character, e.g., C:\
Public Const gstrSEP_DIRALT$ = "/"                      ' Alternate directory separator character
Public Const gstrSEP_EXT$ = "."                         ' Filename extension separator character
Public Const gstrSEP_URLDIR$ = "/"                      ' Separator for dividing directories in URL addresses.

Public Const gstrCOLON$ = ":"
Public Const gstrSwitchPrefix2 = "/"
Public Const gstrCOMMA$ = ","
Public Const gstrDECIMAL$ = "."
Public Const gstrQUOTE$ = """"
Public Const gstrASSIGN$ = "="
Public Const gstrINI_PROTOCOL = "Protocol"

Public Const gstrINIT_POPUP As String = "Initializing...."

Public Const gstrLINE_SEP                     As String = "---------------------------------------------------------------------------------------------------"


'This should remain uppercase
Public Const gstrDCOM = "DCOM"

Public Const gintMAX_SIZE% = 255                        'Maximum buffer size
Public Const gintMAX_PATH_LEN% = 260                    ' Maximum allowed path length including path, filename,
                                                        ' and command line arguments for NT (Intel) and Win95.

Public Const intDRIVE_REMOVABLE% = 2                    'Constants for GetDriveType
Public Const intDRIVE_FIXED% = 3
Public Const intDRIVE_REMOTE% = 4
Public Const intDRIVE_CDROM% = 5

Public Const gintNOVERINFO% = 32767                     'flag indicating no version info


Public Const gsZERO As String = "0"

'MsgError() Constants
Public Const MSGERR_ERROR = 1
Public Const MSGERR_WARNING = 2

'Shell Constants
Public Const NORMAL_PRIORITY_CLASS      As Long = &H20&
Public Const INFINITE                   As Long = -1&

Public Const STATUS_WAIT_0              As Long = &H0
Public Const WAIT_OBJECT_0              As Long = STATUS_WAIT_0

'GetLocaleInfo constants
Public Const LOCALE_FONTSIGNATURE = &H58&           ' font signature

Public Const TCI_SRCFONTSIG = 3

Public Const LANG_CHINESE = &H4
Public Const SUBLANG_CHINESE_TRADITIONAL = &H1           ' Chinese (Taiwan)
Public Const SUBLANG_CHINESE_SIMPLIFIED = &H2            ' Chinese (PR China)
Public Const CHARSET_CHINESESIMPLIFIED = 134
Public Const CHARSET_CHINESEBIG5 = 136

Public Const LANG_JAPANESE = &H11
Public Const CHARSET_SHIFTJIS = 128

Public Const LANG_KOREAN = &H12
Public Const SUBLANG_KOREAN = &H1                        ' Korean (Extended Wansung)
Public Const SUBLANG_KOREAN_JOHAB = &H2                  ' Korean (Johab)
Public Const CHARSET_HANGEUL = 129

Public DoNotShowMessage As Boolean


Public Sub SetFileAttrib(filespec, m_intAttribute As Integer)
'  Dim fso         As Object ', fldr, fil, fil1 'As Folder
'  Dim f           As Object
'  Dim fldr        As Variant
  
  On Error GoTo ErrHandler
  
  Set m_objFileSystemObject = CreateObject("Scripting.FileSystemObject")
  Set m_objTextFile = m_objFileSystemObject.GetFile(filespec)
  
  m_objTextFile.Attributes = m_intAttribute
  
  Set m_objTextFile = Nothing
  
  FILE_NOT_FOUND = False
  
ErrHandler:
   Select Case Err.Number
      Case 53
         FILE_NOT_FOUND = True
                
         Err.Clear
      
         Exit Sub
   
   End Select
End Sub


Public Sub gFolderAttributes(folderspec, m_intAttribute As Integer)
'  Dim fso         As Object ', fldr, fil, fil1 'As Folder
'  Dim f           As Object
'  Dim fldr        As Variant
  
  Set m_objFileSystemObject = CreateObject("Scripting.FileSystemObject")
  Set m_objFolder = m_objFileSystemObject.GetFolder(folderspec)
  
  m_objFolder.Attributes = m_intAttribute
  
  Set m_objFolder = Nothing
  
End Sub


Public Sub MakeDir(ByVal folder As Variant)
  Dim fso         As Object
  Dim f           As Object
  Dim fldr        As Variant
      
  On Error Resume Next
  Set m_objFileSystemObject = CreateObject("Scripting.FileSystemObject")
  Set m_objFolder = m_objFileSystemObject.CreateFolder(folder)
  
  Set m_objFolder = Nothing
  
End Sub


Function gHiddenFolder(ByVal folder As String)
  Dim fso         As Object ', fldr, fil, fil1 'As Folder
  Dim f           As Object
  'Dim fldr        As Variant
  
  'fldr = App.Path & "\System"
  
  Set m_objFileSystemObject = CreateObject("Scripting.FileSystemObject")
  Set m_objFolder = m_objFileSystemObject.CreateFolder(folder)
  
  m_objFolder.Attributes = FILE_ATTRIBUTE_HIDDEN
  
  Set m_objFolder = Nothing
  
End Function


Public Function bFolderExists(ByVal folder As String) As Boolean
  ' Declare variables
  Dim tmpStr As String
  
  'ShowHourglass
  
  Set m_objFileSystemObject = CreateObject("Scripting.FileSystemObject")

  'Check if file exists
  If m_objFileSystemObject.FolderExists(folder) = True Then
    tmpStr = vbCrLf & "Directory exists....."
    'Debug.Print tmpStr
    
    'ShowDefaultPointer
    
    bFolderExists = True
    'Unload Me
    
  Else
    tmpStr = vbCrLf & "Directory does not exist. Create a new directory."
    'Debug.Print tmpStr
    
    'ShowDefaultPointer
    
    bFolderExists = False
    
  End If

End Function


Public Function bFileExists(ByVal TextFile As String) As Boolean
  Dim sTmp As String
  
  ShowHourglass
  
  Set m_objFileSystemObject = CreateObject("Scripting.FileSystemObject")

  'Check if file exists
  If m_objFileSystemObject.FileExists(TextFile) = True Then
    sTmp = vbCrLf & TextFile & " exists."
    'Debug.Print sTmp
    
    'ShowDefaultPointer
    
    bFileExists = True
    'Unload Me
    
  Else
    sTmp = vbCrLf & TextFile & " does not exist."
    'Debug.Print sTmp
    
    'ShowDefaultPointer
    
    bFileExists = False
    
  End If

End Function

Public Sub CopyTextFile(ByVal Source As Variant, destination As Variant)
     
   ' Bypass errors
   On Error GoTo TextFile_ErrHandler
        
   ' Create File System Object
   Set m_objFileSystemObject = CreateObject("Scripting.FileSystemObject")
   Set m_objTextFile = m_objFileSystemObject.CopyFile(Source, destination, True)
   
   Exit Sub
  
TextFile_ErrHandler:
   Select Case Err.Number
      Case 424    ' Object required...
         'Debug.Print Err.Number & vbCrLf
         'Debug.Print Err.Description & vbCrLf
         
         Err.Clear
         
         Exit Sub
      
      Case Else   ' Unknown errors
         'Debug.Print Err.Number & vbCrLf
         'Debug.Print Err.Description & vbCrLf
         
         Err.Clear
         
         Exit Sub
   End Select
End Sub

Public Sub CopyFolderObject(ByVal Source As Variant, destination As Variant)
   On Error GoTo FolderObj_ErrHandler
   Set m_objFileSystemObject = CreateObject("Scripting.FileSystemObject")
   Set m_objFolder = m_objFileSystemObject.CopyFolder(Source, destination)
  
   Exit Sub
  
FolderObj_ErrHandler:
   Select Case Err.Number
      Case 424    ' Object required...
         'Debug.Print Err.Number & vbCrLf
         'Debug.Print Err.Description & vbCrLf
         
         Err.Clear
         
         Exit Sub
      
      Case Else   ' Unknown errors
         'Debug.Print Err.Number & vbCrLf
         'Debug.Print Err.Description & vbCrLf
         
         Err.Clear
         
         Exit Sub
   End Select
  
End Sub

Public Sub SetClearArchiveBit(filespec)
    Dim fs, f, r
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(fs.GetFileName(filespec))
    If f.Attributes And 4 Then
        r = MsgBox("The Archive bit is set, do you want to clear it?", vbYesNo, "Set/Clear Archive Bit")
        If r = vbYes Then
            f.Attributes = f.Attributes - 32
            MsgBox "Archive bit is cleared."
        Else
            MsgBox "Archive bit remains set."
        End If
    Else
        r = MsgBox("The Archive bit is not set. Do you want to set it?", vbYesNo, "Set/Clear Archive Bit")
        If r = vbYes Then
            f.Attributes = f.Attributes + 32
            MsgBox "Archive bit is set."
        Else
            MsgBox "Archive bit remains clear."
        End If
    End If
End Sub

Public Function bHiddenOrReadOnlyFile(ByVal PathName As String) As Boolean
   Dim fs, f, r
   
   
   
GetFileSystemObject:
   
   On Error GoTo ErrHandler
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile(PathName)
   
   GoTo CheckFileAttrib
   
CheckFileAttrib:
   On Error GoTo ErrHandler
   
   Set f = Nothing
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile(PathName)
   
   If f.Attributes And FILE_ATTRIBUTE_READONLY Then
      f.Attributes = FILE_ATTRIBUTE_COMPRESSED + FILE_ATTRIBUTE_HIDDEN + FILE_ATTRIBUTE_READONLY
      
      Set f = Nothing
      
      bHiddenOrReadOnlyFile = True
      
      Exit Function
   
   ElseIf f.Attributes And FILE_ATTRIBUTE_HIDDEN Then
      f.Attributes = FILE_ATTRIBUTE_COMPRESSED + FILE_ATTRIBUTE_HIDDEN + FILE_ATTRIBUTE_READONLY
      
      Set f = Nothing
      
      bHiddenOrReadOnlyFile = True
      
      Exit Function
   
   Else
      f.Attributes = FILE_ATTRIBUTE_COMPRESSED + FILE_ATTRIBUTE_HIDDEN + FILE_ATTRIBUTE_READONLY
      
      bHiddenOrReadOnlyFile = False
      
      Set f = Nothing
            
      Exit Function
      
   End If

ErrHandler:
   
   Select Case Err.Number
      Case 53
         Err.Clear
            
         SetFileAttrib PathName, FILE_ATTRIBUTE_NORMAL
            
         If FILE_NOT_FOUND Then
            bHiddenOrReadOnlyFile = False
            
            Set f = Nothing
            
            Exit Function
            
         End If
         
         GoTo CheckFileAttrib
         
         Exit Function
      
      Case 424
         Err.Clear
            
         SetFileAttrib PathName, FILE_ATTRIBUTE_NORMAL
         
         GoTo CheckFileAttrib
         
         Exit Function
      
   End Select
   
End Function

Public Function bSystemFolder(ByVal PathName As String) As Boolean
   Dim fs, f, r
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFolder(PathName)
   
   If f.Attributes = 4 Then
      Set f = Nothing
      
      bSystemFolder = True
      
      Exit Function
   
   Else
      bSystemFolder = False
      
      Set f = Nothing
            
      Exit Function
      
   End If

End Function



Public Sub DeleteTextFile(ByVal filespec As Variant)
   '// Bypass errors
   On Local Error GoTo ErrHandler
        
   '// Create File System Object
   Set m_objFileSystemObject = CreateObject("Scripting.FileSystemObject")
   Set m_objTextFile = m_objFileSystemObject.DeleteFile(filespec, True)
   
   Exit Sub
  
ErrHandler:
   Err.Clear
   
   Exit Sub
   
End Sub


Public Sub xDeleteFolder(ByVal folderspec As Variant)
   Dim f, fs
   
   '// Bypass errors
   On Error GoTo ErrHandler
        
   '// Create File System Object
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.DeleteFolder(folderspec, True)
   
   Exit Sub
  
ErrHandler:
   If Err.Number <> 0 Then
      Err.Clear
   End If
   
   Exit Sub
   
End Sub




'-----------------------------------------------------------
' FUNCTION: DirExists
'
' Determines whether the specified directory name exists.
' This function is used (for example) to determine whether
' an installation floppy is in the drive by passing in
' something like 'A:\'.
'
' IN: [strDirName] - name of directory to check for
'
' Returns: True if the directory exists, False otherwise
'-----------------------------------------------------------
'
Public Function DirExists(ByVal strDirName As String) As Boolean
    On Error Resume Next

    DirExists = (GetAttr(strDirName) And vbDirectory) = vbDirectory

    Err.Clear
End Function



'-----------------------------------------------------------
' SUB: AddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'
' IN/OUT: [strPathName] - path to add separator to
'-----------------------------------------------------------
'
Public Sub AddDirSep(strPathName As String)
    strPathName = RTrim$(strPathName)
    If Right$(strPathName, Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR Then
        If Right$(strPathName, Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
            strPathName = strPathName & gstrSEP_DIR
        End If
    End If
End Sub

'-----------------------------------------------------------
' SUB: RemoveDirSep
' Removes a trailing directory path separator (back slash)
' at the end of a pathname if one exists
'
' IN/OUT: [strPathName] - path to remove separator from
'-----------------------------------------------------------
'
Public Sub RemoveDirSep(strPathName As String)
    Select Case Right$(strPathName, 1)
    Case gstrSEP_DIR, gstrSEP_DIRALT
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End Select
End Sub

'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Public Function FileExists(ByVal strPathName As String) As Boolean
   Dim intFileNum As Integer
   
   On Error Resume Next
   
   '
   ' If the string is quoted, remove the quotes.
   '
   strPathName = strUnQuoteString(strPathName)
   '
   'Remove any trailing directory separator character
   '
   If Right$(strPathName, 1) = gstrSEP_DIR Then
      strPathName = Left$(strPathName, Len(strPathName) - 1)
   End If
   
   '
   'Attempt to open the file, return value of this function is False
   'if an error occurs on open, True otherwise
   '
   intFileNum = FreeFile
   Open strPathName For Input As intFileNum
   
   FileExists = (Err.Number = 0)
   
   Close intFileNum
   
   Err.Clear
   
End Function




Public Function strQuoteString(strUnQuotedString As String, Optional vForce As Boolean = False, Optional vTrim As Boolean = True)
'
' This routine adds quotation marks around an unquoted string, by default.  If the string is already quoted
' it returns without making any changes unless vForce is set to True (vForce defaults to False) except that white
' space before and after the quotes will be removed unless vTrim is False.  If the string contains leading or
' trailing white space it is trimmed unless vTrim is set to False (vTrim defaults to True).
'
    Dim strQuotedString As String

    strQuotedString = strUnQuotedString
    '
    ' Trim$ the string if necessary
    '
    If vTrim Then
        strQuotedString = Trim$(strQuotedString)
    End If
    '
    ' See if the string is already quoted
    '
    If Not vForce Then
        If Left$(strQuotedString, 1) = gstrQUOTE Then
            If Right$(strQuotedString, 1) = gstrQUOTE Then
                '
                ' String is already quoted.  We are done.
                '
                GoTo DoneQuoteString
            End If
        End If
    End If
    '
    ' Add the quotes
    '
    strQuotedString = gstrQUOTE & strQuotedString & gstrQUOTE
    
DoneQuoteString:
    strQuoteString = strQuotedString
End Function

Public Function strUnQuoteString(ByVal strQuotedString As String)
'
' This routine tests to see if strQuotedString is wrapped in quotation
' marks, and, if so, remove them.
'
    strQuotedString = Trim$(strQuotedString)

    If Mid$(strQuotedString, 1, 1) = gstrQUOTE Then
        If Right$(strQuotedString, 1) = gstrQUOTE Then
            '
            ' It's quoted.  Get rid of the quotes.
            '
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function

Public Function StringFromBuffer(Buffer As String) As String
    Dim nPos As Long

    nPos = InStr(Buffer, vbNullChar)
    If nPos > 0 Then
        StringFromBuffer = Left$(Buffer, nPos - 1)
    Else
        StringFromBuffer = Buffer
    End If
End Function


Public Function GetFileExtension(filespec As Variant) As String
   Dim fso, fs, f
   
'   Debug.Print "Pathname = "; filespec
   
   If Len(filespec) = 0 Then
      GetFileExtension = ""
      Exit Function
   End If
   
   '// Bypass errors
   On Local Error GoTo ErrHandler
        
   '// Create File System Object
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(fso.GetFileName(filespec))
   
'   Debug.Print f.Name
   
   GetFileExtension = fso.GetExtensionName(filespec)
   
   GetFileExtension = "." & LCase(GetFileExtension)
   
'   Debug.Print GetFileExtension
   
   Exit Function
  
ErrHandler:
   Err.Clear
   
   GetFileExtension = ""
   
   Exit Function
   
End Function


