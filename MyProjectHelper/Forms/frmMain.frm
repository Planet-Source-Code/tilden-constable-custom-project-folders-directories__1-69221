VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpenFolder 
      Caption         =   "Open Folder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5700
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame fraLine 
      Height          =   90
      Index           =   0
      Left            =   90
      TabIndex        =   24
      Top             =   5610
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   26
      Top             =   5700
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6900
      TabIndex        =   27
      Top             =   5700
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   5565
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   7815
      Begin VB.TextBox txtNotes 
         Height          =   315
         Left            =   1230
         MaxLength       =   255
         TabIndex        =   16
         Text            =   "Notes, Remarks and Comments"
         Top             =   5010
         Width           =   3255
      End
      Begin VB.CommandButton cmdCheckAll 
         Caption         =   "Check All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6210
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1245
      End
      Begin VB.CommandButton cmdUncheck 
         Caption         =   "Uncheck All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4860
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1245
      End
      Begin VB.Frame fraLine 
         Height          =   90
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   7155
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Default"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   930
         Width           =   1275
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6210
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   930
         Width           =   1275
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4830
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   930
         Width           =   1275
      End
      Begin VB.TextBox txtPathname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   510
         Width           =   7125
      End
      Begin VB.CheckBox chkSubFolder 
         Caption         =   "Other"
         Height          =   195
         Index           =   3
         Left            =   5490
         TabIndex        =   23
         Top             =   5070
         Width           =   1515
      End
      Begin VB.CheckBox chkSubFolder 
         Caption         =   "Jpegs"
         Height          =   195
         Index           =   2
         Left            =   5490
         TabIndex        =   22
         Top             =   4560
         Width           =   1515
      End
      Begin VB.CheckBox chkSubFolder 
         Caption         =   "Gifs"
         Height          =   195
         Index           =   1
         Left            =   5490
         TabIndex        =   21
         Top             =   4020
         Width           =   1515
      End
      Begin VB.CheckBox chkSubFolder 
         Caption         =   "Bitmaps"
         Height          =   195
         Index           =   0
         Left            =   5490
         TabIndex        =   20
         Top             =   3510
         Width           =   1515
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "My Pictures"
         Height          =   195
         Index           =   8
         Left            =   4620
         TabIndex        =   19
         Top             =   2970
         Width           =   1515
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "Icons"
         Height          =   195
         Index           =   7
         Left            =   4620
         TabIndex        =   18
         Top             =   2460
         Width           =   1515
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "Cursors"
         Height          =   195
         Index           =   6
         Left            =   4620
         TabIndex        =   17
         Top             =   1950
         Width           =   1515
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "Notes, Remarks and Comments"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   15
         Top             =   5070
         Width           =   585
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "User Controls"
         Height          =   195
         Index           =   4
         Left            =   930
         TabIndex        =   14
         Top             =   4020
         Width           =   1515
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "Designers"
         Height          =   195
         Index           =   3
         Left            =   930
         TabIndex        =   13
         Top             =   3510
         Width           =   1515
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "Classes"
         Height          =   195
         Index           =   2
         Left            =   930
         TabIndex        =   12
         Top             =   2970
         Width           =   1515
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "Modules"
         Height          =   195
         Index           =   1
         Left            =   930
         TabIndex        =   11
         Top             =   2460
         Width           =   1515
      End
      Begin VB.CheckBox chkFolder 
         Caption         =   "Forms"
         Height          =   195
         Index           =   0
         Left            =   930
         TabIndex        =   10
         Top             =   1950
         Width           =   1515
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Select one or more of the options (sub folders) below."
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   1530
         Width           =   3750
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Browse/enter path of Project Directory"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   2715
      End
      Begin VB.Image imgIcon1 
         Height          =   480
         Index           =   3
         Left            =   4830
         Picture         =   "frmMain.frx":27A2
         Tag             =   "202.ico"
         Top             =   4890
         Width           =   480
      End
      Begin VB.Image imgIcon1 
         Height          =   480
         Index           =   2
         Left            =   4830
         Picture         =   "frmMain.frx":4B24
         Tag             =   "202.ico"
         Top             =   4410
         Width           =   480
      End
      Begin VB.Image imgIcon1 
         Height          =   480
         Index           =   1
         Left            =   4830
         Picture         =   "frmMain.frx":6EA6
         Tag             =   "202.ico"
         Top             =   3900
         Width           =   480
      End
      Begin VB.Image imgIcon1 
         Height          =   480
         Index           =   0
         Left            =   4830
         Picture         =   "frmMain.frx":9228
         Tag             =   "202.ico"
         Top             =   3390
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   8
         Left            =   3900
         Picture         =   "frmMain.frx":B5AA
         Tag             =   "204.ico"
         Top             =   2850
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   7
         Left            =   3900
         Picture         =   "frmMain.frx":CF3C
         Tag             =   "205.ico"
         Top             =   2340
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   6
         Left            =   3900
         Picture         =   "frmMain.frx":F6DE
         Tag             =   "205.ico"
         Top             =   1830
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   5
         Left            =   300
         Picture         =   "frmMain.frx":11E80
         Tag             =   "205.ico"
         Top             =   4890
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   4
         Left            =   270
         Picture         =   "frmMain.frx":14622
         Tag             =   "202.ico"
         Top             =   3900
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   3
         Left            =   270
         Picture         =   "frmMain.frx":169A4
         Tag             =   "202.ico"
         Top             =   3390
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   2
         Left            =   270
         Picture         =   "frmMain.frx":18D26
         Tag             =   "202.ico"
         Top             =   2850
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   270
         Picture         =   "frmMain.frx":1B0A8
         Tag             =   "202.ico"
         Top             =   2340
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   270
         Picture         =   "frmMain.frx":1D42A
         Tag             =   "202.ico"
         Top             =   1830
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarFilename                            As String
Private mvarDefault                             As String
Private mvarRootIcon                            As String

Private Const mvarGeneral                       As String = "Notes, Remarks and Comments"

' The Broadcast event has one argument,
'   the message to be sent.  The argument
'   is ByRef, so recipients can change it.
Event Broadcast(Message As String)

Private Sub chkFolder_Click(Index As Integer)
   Dim itmX                                     As Integer
   
   Select Case Index
      Case Is = 5
         Select Case chkFolder(Index).Value
            Case vbChecked
               txtNotes.Enabled = True
            Case vbUnchecked
               txtNotes.Enabled = False
         End Select
      
      Case Is = 8
         Select Case chkFolder(Index).Value
            Case vbChecked
               For itmX = 0 To (chkSubFolder.Count - 1)
                  chkSubFolder(itmX).Enabled = True
               Next
               
            Case vbUnchecked
               For itmX = 0 To (chkSubFolder.Count - 1)
                  With chkSubFolder(itmX)
                     .Value = vbUnchecked
                     .Enabled = False
                  End With
               Next
            
         End Select
         
      Case Else
         Exit Sub
         
   End Select
End Sub

Private Sub chkFolder_KeyPress(Index As Integer, KeyAscii As Integer)
   On Error Resume Next
   
   Select Case KeyAscii
      Case vbKeyReturn
         SendKeys vbTab
         KeyAscii = 0
   
   End Select
   
End Sub

Private Sub chkSubFolder_KeyPress(Index As Integer, KeyAscii As Integer)
   On Error Resume Next
   
   Select Case KeyAscii
      Case vbKeyReturn
         SendKeys vbTab
         KeyAscii = 0
   
   End Select

End Sub

Private Sub cmdBrowse_Click()
   Dim varPathname As String
   
   varPathname = ""
   
   varPathname = BrowseForFolder("C:", Me.hWnd, "Create/select a Project Folder")
   
   If Len(varPathname) = 0 Then Exit Sub
   
   txtPathname = varPathname
   
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   
End Sub

Private Sub Initialize()
   Dim ctl                                      As Control
   
   NewDoEvents
   
   mvarFilename = App.Path
   mvarDefault = "C:\My Project"
   mvarRootIcon = App.Path & "\201.ico"
   
   For Each ctl In Me.Controls
      If TypeOf ctl Is CommandButton Or _
         TypeOf ctl Is Frame Or _
         TypeOf ctl Is TextBox Or _
         TypeOf ctl Is Label Or _
         TypeOf ctl Is CheckBox Then
         
         ctl.FontName = "Tahoma"
         
      End If
   Next
   
   txtPathname = mvarDefault
   txtNotes = mvarGeneral
   chkFolder(5).Caption = mvarGeneral
   
   txtNotes_Change
   
   cmdCheckAll_Click
   
End Sub

Private Sub cmdCheckAll_Click()
   Dim ctl As Control
   
   For Each ctl In Me.Controls
      If TypeOf ctl Is CheckBox Then
         ctl.Value = vbChecked
      End If
   Next
   
   txtNotes.Enabled = True
   
End Sub

Private Sub cmdClear_Click()
   txtPathname = vbNullString
   
   ControlsEnabled False
   
End Sub

Private Sub cmdDefault_Click()
   txtPathname = mvarDefault
End Sub

Private Sub cmdOK_Click()
   Dim varPathname                              As String
   Dim MsgStr                                   As String
   Dim lRet
   
   varPathname = ""
   
   varPathname = CStr(txtPathname)
   
   ConfgProjectDirectories varPathname
      
   If Len(varPathname) = 0 Then Exit Sub
   
   MsgStr = ""
   
   MsgStr = App.ProductName & " has successfully completed setting up your Project directory." ' in '" & varPathname & "'."

   MsgBox MsgStr, vbInformation
   
   OpenProjectFolderEx varPathname
   
   'lRet = BrowsePath(Me.hWnd, varPathname)
   
End Sub

Private Sub cmdOpenFolder_Click()
   Dim varPathname                              As String
   
   varPathname = ""

   varPathname = CStr(txtPathname)
   
   If Not DirExists(varPathname) Then Exit Sub

   OpenFolderEx Me.hWnd, varPathname
   
End Sub

Private Sub cmdUncheck_Click()
   Dim ctl As Control
   
   For Each ctl In Me.Controls
      If TypeOf ctl Is CheckBox Then
         ctl.Value = vbUnchecked
      End If
   Next
   
   txtNotes.Enabled = False
   
End Sub


Private Sub Form_Load()
   NewDoEvents
   
   Initialize
   
   NewDoEvents
   
End Sub

Sub ControlsEnabled(Optional Enabled As Boolean = False)
   Dim ctl As Control
   
   cmdClear.Enabled = Enabled
   cmdOK.Enabled = Enabled
   
   cmdUncheck.Enabled = Enabled
   cmdCheckAll.Enabled = Enabled
   
   For Each ctl In Me.Controls
      If TypeOf ctl Is CheckBox Then
         ctl.Enabled = Enabled
      End If
   Next
   
   txtNotes.Enabled = Enabled
'   txtNotes.Locked = True
   
End Sub

Private Sub txtNotes_Change()
   Dim WrkStr As String
   
   WrkStr = CStr(txtNotes)
   
   '
   ' Raise the Broadcast event.  Note that
   '   there's no way of knowing if there
   '   are any receivers handling the
   '   event.
   RaiseEvent Broadcast(WrkStr)
   
   '
   ' Display the message after all
   '   receivers (if any) have handled
   '   it.  Note that there's no way
   '   to know which receiver altered the
   '   message, or what interim values
   '   the message may have had.
'   Label1 = WrkStr
   chkFolder(5).Caption = WrkStr
   
End Sub

Private Sub txtNotes_GotFocus()
   SelectEx txtNotes
   
End Sub

Private Sub txtNotes_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   
   Select Case KeyAscii
      Case vbKeyReturn
         SendKeys vbTab
         KeyAscii = 0
   
   End Select
   
End Sub

Private Sub txtNotes_LostFocus()
   UnselectEx txtNotes
   
End Sub

Private Sub txtNotes_Validate(Cancel As Boolean)
   If Len(txtNotes) = 0 Then
      txtNotes = mvarGeneral
   End If
   
End Sub

Private Sub txtPathname_Change()

   Select Case Len(txtPathname)
      Case Is = 0
         ControlsEnabled False
      Case Is <> 0
         ControlsEnabled True
   End Select
   
End Sub

Private Sub txtPathname_GotFocus()
   SelectEx txtPathname
End Sub

Private Sub txtPathname_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   
   Select Case KeyAscii
      Case vbKeyReturn
         SendKeys vbTab
         KeyAscii = 0
   End Select
   
End Sub

Private Sub txtPathname_LostFocus()
   UnselectEx txtPathname
   
End Sub

Sub ConfgProjectDirectories(Optional lpPathname As String = vbNullString)
   Dim varFilename                              As String
   Dim varDstFilename                           As String
   Dim varSrcFilename                           As String
   Dim varPathname                              As String
   
   Dim itmX                                     As Integer
   
   If Len(lpPathname) = 0 Then Exit Sub
   
   varPathname = lpPathname & gstrSEP_DIR
   varSrcFilename = mvarRootIcon
   varDstFilename = lpPathname & gstrSEP_DIR & "201.ico"
   
   If DirExists(lpPathname) Then
      If Not FileExists(varDstFilename) Then
         CopyTextFile varSrcFilename, varPathname
      End If
      
      SetAttr lpPathname, vbNormal
      
      WriteConfigFile lpPathname, varDstFilename
   
      SetAttr lpPathname, vbSystem + vbReadOnly
      SetAttr varDstFilename, vbHidden
      
      NewDoEvents
      
   Else
      MkDir lpPathname
         
      NewDoEvents
      
      If Not FileExists(varDstFilename) Then
         CopyTextFile varSrcFilename, varPathname
      End If
            
      SetAttr lpPathname, vbNormal
      
      WriteConfigFile lpPathname, varDstFilename
   
      SetAttr lpPathname, vbSystem + vbReadOnly
      SetAttr varDstFilename, vbHidden
      
      NewDoEvents
            
   End If
   
   For itmX = 0 To chkFolder.Count - 1
      If chkFolder(itmX).Value = vbChecked Then
         varPathname = "": varDstFilename = ""
         
         varPathname = lpPathname & gstrSEP_DIR & chkFolder(itmX).Caption
         
         varSrcFilename = App.Path & gstrSEP_DIR & imgIcon(itmX).Tag
         varDstFilename = varPathname & gstrSEP_DIR & imgIcon(itmX).Tag
         
         If DirExists(varPathname) Then
            If Not FileExists(varDstFilename) Then
               CopyTextFile varSrcFilename, varPathname & gstrSEP_DIR
            End If
            
            SetAttr varPathname, vbNormal
            
            WriteConfigFile varPathname, varDstFilename
         
            SetAttr varPathname, vbSystem + vbReadOnly
            
            SetAttr varDstFilename, vbHidden
            
            NewDoEvents
            
         Else
            MkDir varPathname
               
            NewDoEvents
            
            If Not FileExists(varDstFilename) Then
               CopyTextFile varSrcFilename, varPathname & gstrSEP_DIR
            End If
                  
            SetAttr varPathname, vbNormal
            
            WriteConfigFile varPathname, varDstFilename
         
            SetAttr varPathname, vbSystem + vbReadOnly
            
            SetAttr varDstFilename, vbHidden
            
            NewDoEvents
                  
         End If
      
      End If
   Next
   
   If chkFolder(8).Value = vbChecked Then
      For itmX = 0 To chkSubFolder.Count - 1
         If chkSubFolder(itmX).Value = vbChecked Then
            varPathname = "": varDstFilename = ""
            
            varPathname = lpPathname & gstrSEP_DIR & chkFolder(8).Caption & gstrSEP_DIR & chkSubFolder(itmX).Caption
            
            varSrcFilename = App.Path & gstrSEP_DIR & imgIcon1(itmX).Tag
            varDstFilename = varPathname & gstrSEP_DIR & imgIcon1(itmX).Tag
            
            If DirExists(varPathname) Then
               If Not FileExists(varDstFilename) Then
                  CopyTextFile varSrcFilename, varPathname & gstrSEP_DIR
               End If
               
               SetAttr varPathname, vbNormal
               
               WriteConfigFile varPathname, varDstFilename
            
               SetAttr varPathname, vbSystem + vbReadOnly
               
               SetAttr varDstFilename, vbHidden
               
               NewDoEvents
               
            Else
               MkDir varPathname
                  
               NewDoEvents
               
               If Not FileExists(varDstFilename) Then
                  CopyTextFile varSrcFilename, varPathname & gstrSEP_DIR
               End If
                     
               SetAttr varPathname, vbNormal
               
               WriteConfigFile varPathname, varDstFilename
            
               SetAttr varPathname, vbSystem + vbReadOnly
               
               SetAttr varDstFilename, vbHidden
               
               NewDoEvents
                     
            End If
         
         End If
      Next
   End If
   
End Sub
