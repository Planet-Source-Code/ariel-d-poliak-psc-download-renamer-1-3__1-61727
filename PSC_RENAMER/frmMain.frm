VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSC Download FileName Fixer - Based on Flyhole"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Tag             =   "!"
   Begin VB.Frame fraCreditsHolder 
      Height          =   1215
      Left            =   7440
      TabIndex        =   29
      Top             =   2280
      Width           =   2055
      Begin VB.Label lblCredits3 
         AutoSize        =   -1  'True
         Caption         =   "by Crock."
         Height          =   195
         Left            =   1320
         TabIndex        =   32
         Top             =   960
         Width           =   675
      End
      Begin VB.Label lblCredits2 
         AutoSize        =   -1  'True
         Caption         =   "Flyhole"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   840
         MouseIcon       =   "frmMain.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblCredits1 
         AutoSize        =   -1  'True
         Caption         =   "Based on "
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame fraSFDetails 
      Caption         =   "Selected File Details"
      Height          =   1575
      Left            =   7440
      TabIndex        =   22
      Top             =   120
      Width           =   2055
      Begin VB.PictureBox picCFXPBugFixfrmMain 
         BorderStyle     =   0  'None
         Height          =   1305
         Index           =   3
         Left            =   120
         ScaleHeight     =   1305
         ScaleWidth      =   1860
         TabIndex        =   23
         Top             =   240
         Width           =   1855
         Begin VB.Frame fraProjName 
            Caption         =   "Project Name"
            Height          =   1215
            Left            =   0
            TabIndex        =   24
            Top             =   -18
            Width           =   1815
            Begin VB.TextBox txtProjectName 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   25
               Top             =   240
               Width           =   1575
            End
         End
      End
   End
   Begin VB.Frame fraSelectedReadme 
      Caption         =   "Sele&cted README"
      Height          =   735
      Left            =   4080
      TabIndex        =   19
      Tag             =   "!"
      Top             =   1800
      Width           =   3255
      Begin VB.TextBox txtSelectedReadme 
         Height          =   285
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Tag             =   "!"
         Top             =   1065
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewHide 
         Caption         =   "&View"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Tag             =   "!"
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Tag             =   "!"
         Top             =   825
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Frame fraLocation 
      Caption         =   "Location"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.PictureBox picCFXPBugFixfrmMain 
         BorderStyle     =   0  'None
         Height          =   5445
         Index           =   0
         Left            =   100
         ScaleHeight     =   5445
         ScaleWidth      =   3660
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   276
         Width           =   3655
         Begin VB.Frame fraAbsLocation 
            Caption         =   "&Absolute Location"
            Height          =   615
            Left            =   0
            TabIndex        =   7
            Top             =   4680
            Width           =   3615
            Begin VB.TextBox txtAbsLocation 
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   240
               Width           =   3375
            End
         End
         Begin VB.Frame fraDrive 
            Caption         =   "&Drive"
            Height          =   615
            Left            =   20
            TabIndex        =   1
            Top             =   -18
            Width           =   3615
            Begin VB.DriveListBox drvDrive 
               Height          =   315
               Left            =   120
               TabIndex        =   2
               Top             =   240
               Width           =   3375
            End
         End
         Begin VB.Frame fraDirectory 
            Caption         =   "D&irectory"
            Height          =   3135
            Left            =   20
            TabIndex        =   3
            Top             =   702
            Width           =   3615
            Begin VB.DirListBox dirDirectory 
               Height          =   2790
               Left            =   120
               TabIndex        =   4
               Top             =   240
               Width           =   3375
            End
         End
         Begin VB.Frame fraFilter 
            Caption         =   "&Filter"
            Height          =   615
            Left            =   20
            TabIndex        =   5
            Top             =   3942
            Width           =   3615
            Begin VB.TextBox txtFilter 
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   6
               Text            =   "*.zip"
               Top             =   240
               Width           =   3375
            End
         End
      End
   End
   Begin VB.Frame fraZips 
      Caption         =   "Files Ma&tching Criteria"
      Height          =   1575
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   3255
      Begin VB.FileListBox filZips 
         Height          =   1260
         Hidden          =   -1  'True
         Left            =   120
         MultiSelect     =   2  'Extended
         Pattern         =   "*.zip"
         System          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fraToRename 
      Caption         =   "Files T&o Rename"
      Height          =   1575
      Left            =   4080
      TabIndex        =   14
      Top             =   4320
      Width           =   4575
      Begin VB.ListBox lstFilesToRename 
         Height          =   1230
         ItemData        =   "frmMain.frx":0BD4
         Left            =   120
         List            =   "frmMain.frx":0BD6
         MultiSelect     =   2  'Extended
         TabIndex        =   15
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame fraListOps 
      Caption         =   "Li&st Operations"
      Height          =   1575
      Left            =   4080
      TabIndex        =   11
      Top             =   2640
      Width           =   3015
      Begin VB.PictureBox picCFXPBugFixfrmMain 
         BorderStyle     =   0  'None
         Height          =   1245
         Index           =   1
         Left            =   120
         ScaleHeight     =   1245
         ScaleWidth      =   2820
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   2820
         Begin VB.CommandButton cmdRemoveAll 
            Caption         =   "Re&move All"
            Height          =   495
            Left            =   1440
            TabIndex        =   17
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdRemoveSelected 
            Caption         =   "&Remove Selected"
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "Add A&ll"
            Height          =   495
            Left            =   1440
            TabIndex        =   13
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddSelected 
            Caption         =   "Add S&elected"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "DO IT&!"
      Height          =   1455
      Left            =   8760
      TabIndex        =   18
      Top             =   4440
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FRAMEWORK by Crock
Option Explicit
Private ZipName      As String
Private FilePath     As String
Private ProgName     As String
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub cmdAddAll_Click()

  Dim x As Integer

  For x = 0 To filZips.ListCount - 1
    lstFilesToRename.AddItem FilePath & filZips.List(x)
  Next

End Sub

Private Sub cmdAddSelected_Click()

  Dim x As Integer

  With filZips
    For x = 0 To .ListCount - 1
      If .Selected(x) Then
        lstFilesToRename.AddItem FilePath & .List(x)
      End If
    Next
  End With 'filZips

End Sub

Private Sub cmdGO_Click()

  Dim x           As Integer
  Dim strFilePath As String
  Dim a           As Control

  For x = lstFilesToRename.ListCount - 1 To 0 Step -1
    strFilePath = lstFilesToRename.List(x)
    ZipName = Right$(strFilePath, Len(strFilePath) - InStrRev(strFilePath, "\", , vbTextCompare))
    FilePath = Left$(strFilePath, Len(strFilePath) - Len(ZipName))
    uZip
    ZipName = txtProjectName.Text
    If ZipName <> "NO README FOUND!" Then
      ZipName = SafeFileName(ZipName, "zip")
      SetAttr strFilePath, GetAttr(strFilePath) And Not vbReadOnly  '
      Name strFilePath As FilePath & ZipName
      lstFilesToRename.RemoveItem x
    End If
  Next
  For Each a In Me.Controls
    a.Refresh
  Next
  Select Case lstFilesToRename.ListCount
   Case 0
    MsgBox "DONE!"
   Case Else
    MsgBox "The files left on the list were not renamed. One cause of this is missing PSC_README files. Please manually open the ZIP files and check if a valid @PSC_README*.txt file exists within it."
  End Select

End Sub

Private Sub cmdRemoveAll_Click()

  Dim x As Integer

  With lstFilesToRename
    For x = .ListCount - 1 To 0 Step -1
      .RemoveItem x
      .Refresh
    Next
  End With 'lstFilesToRename

End Sub

Private Sub cmdRemoveSelected_Click()

  Dim x As Integer

  With lstFilesToRename
    For x = .ListCount - 1 To 0 Step -1
      If .Selected(x) Then
        .RemoveItem x
        .Refresh
      End If
    Next
  End With 'lstFilesToRename

End Sub

Private Sub cmdViewHide_Click()

  Dim a              As Control
  Static oldLeft     As Single
  Static oldTop      As Single
  Static oldWidth    As Single
  Static oldHeight   As Single
  Static oldSRWidth  As Single
  Static oldSRHeight As Single

  With fraSelectedReadme
    Select Case cmdViewHide.Caption
     Case "&View"
      oldLeft = .Left
      oldTop = .Top
      oldWidth = .Width
      oldHeight = .Height
      .Left = 0
      .Top = 0
      .Width = Me.ScaleWidth
      .Height = Me.ScaleHeight
      .ZOrder 0
      With txtSelectedReadme
        oldSRWidth = .Width
        oldSRHeight = .Height
        .Width = fraSelectedReadme.Width - .Left - Me.ScaleX(10, vbPixels, vbTwips)
        .Height = fraSelectedReadme.Height - .Top - Me.ScaleY(10, vbPixels, vbTwips)
        .Visible = True
      End With 'txtSelectedReadme
      lblFileName.Visible = True
      For Each a In Me.Controls
        If a.Tag <> "!" Then
          a.Enabled = False
        End If
      Next
      cmdViewHide.Caption = "&Hide"
     Case "&Hide"
      .ZOrder 1
      With txtSelectedReadme
        .Width = oldSRWidth
        .Height = oldSRHeight
        .Visible = False
      End With 'txtSelectedReadme
      lblFileName.Visible = False
      .Height = oldHeight
      .Width = oldWidth
      .Top = oldTop
      .Left = oldLeft
      For Each a In Me.Controls
        a.Enabled = True
      Next
      cmdViewHide.Caption = "&View"
    End Select
  End With 'FRASELECTEDREADME

End Sub

Private Sub dirDirectory_Change()

  Dim strBSlash As String

  If Right$(dirDirectory.Path, 1) <> "\" Then
    strBSlash = "\"
    Else 'NOT RIGHT$(DIRDIRECTORY.PATH,...
    strBSlash = vbNullString
  End If
  txtAbsLocation.Text = dirDirectory.Path & strBSlash & " - " & txtFilter.Text
  filZips.Path = dirDirectory.Path
  If Right$(dirDirectory.Path, 1) = "\" Then
    FilePath = dirDirectory.Path
    Else 'NOT RIGHT$(DIRDIRECTORY.PATH,...
    FilePath = dirDirectory.Path & "\"
  End If

End Sub

Private Sub drvDrive_Change()

  On Error Resume Next
  If Left$(dirDirectory.Path, 1) <> Left$(drvDrive.Drive, 1) Then
    dirDirectory.Path = drvDrive.Drive
  End If
  On Error GoTo 0

End Sub

Private Sub filZips_Click()

  If ZipName = filZips.List(filZips.ListIndex) Then
    Exit Sub    ' No need to continue, as this is the current zip.
    Else 'NOT ZIPNAME...
    txtSelectedReadme.Text = vbNullString ' Clear the text
  End If
  With filZips
    If .ListCount > 0 Then
      ZipName = .List(.ListIndex)   ' Make the selected item the zipname
      ' Put some info in the tool tip
      .ToolTipText = Left$(FileDateTime(FilePath & ZipName), 8) & " " & FileLen(FilePath & ZipName) & " bytes"
      uZip    ' Call the uZip procedure
      Else 'NOT .LISTCOUNT...
      MsgBox "No zip files found in " & FilePath
    End If
  End With 'filZips

End Sub

Private Sub filZips_DblClick()

  RunAssociated FilePath & ZipName, Me.hwnd

End Sub

Private Sub filZips_PathChange()

  txtProjectName.Text = vbNullString

End Sub

Private Sub Form_Initialize()

  Dim strSlash  As String
  Dim strBuffer As String
  Dim intFile   As Integer
  Dim x         As Integer

  If Right$(App.Path, 1) <> "\" Then
    strSlash = "\"
  Else 'NOT RIGHT$(APP.PATH,...
    strSlash = vbNullString
  End If
  If LenB(Dir(App.Path & strSlash & "unzip32.dll")) = 0 Then
    x = SaveResItemToDisk(101, 100, App.Path & strSlash & "unzip32.dll")
    If x <> 0 Then
      MsgBox "Error creating UNZIP32.DLL, a necessary file. This application needs to run from a location which allows Read/Write operations. Contact the author for more details."
    End If
  End If
  InitCommonControls

End Sub

Private Sub Form_Load()

  ' Set path

  If Right$(App.Path, 1) = "\" Then
    FilePath = App.Path
    Else 'NOT RIGHT$(APP.PATH,...
    FilePath = App.Path & "\"
  End If
  drvDrive_Change
  dirDirectory_Change
  txtFilter_Change

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim I As Integer

  'close all sub forms
  For I = Forms.Count - 1 To 1 Step -1
    Unload Forms(I)
  Next

End Sub

Private Sub lblCredits2_Click()

  RunAssociated "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=15059&lngWId=1", Me.hwnd

End Sub

Private Sub txtFilter_Change()

  filZips.Pattern = txtFilter.Text

End Sub

Private Sub uZip()

  Dim strTitle As String
  Dim Crit     As Variant
  Dim I        As Long

  strTitle = "NO README FOUND!"
  Screen.MousePointer = vbHourglass   ' Change mouse pointer to hourglass.
  'lblFileName.Visible = False
  'txtSelectedReadme.Visible = False
  Cls
  txtSelectedReadme.Text = vbNullString
  '-- Init Global Message Variables
  uZipInfo = vbNullString
  uZipNumber = 0   ' Holds The Number Of Zip Files
  ' List of variables from the author of VBUnzBas.
  '-- Public Variables For Setting The UNZIP32.DLL DCLIST Structure
  '-- These Must Be Set Before The Actual Call To VBUnZip32
  'uExtractOnlyNewer = 0  ' 1 = Extract Only Newer/New, Else 0
  'uSpaceUnderScore = 0   ' 1 = Convert Space To Underscore, Else 0
  'uPromptOverWrite = 1   ' 1 = Prompt To Overwrite Required, Else 0
  uQuiet = 2             ' 2 = No Messages, 1 = Less, 0 = All
  uWriteStdOut = 1       ' 1 = Write To Stdout, Else 0
  'uTestZip = 0           ' 1 = Test Zip File, Else 0
  uExtractList = 1       ' 0 = Extract, 1 = List Contents
  'uFreshenExisting = 0   ' 1 = Update Existing by Newer, Else 0
  uDisplayComment = 0    ' 1 = Display Zip File Comment, Else 0
  'uHonorDirectories = 1  ' 1 = Honor Directories, Else 0
  'uOverWriteFiles = 0    ' 1 = Overwrite Files, Else 0
  'uConvertCR_CRLF = 0    ' 1 = Convert CR To CRLF, Else 0
  'uVerbose = 1           ' 1 = Zip Info Verbose
  uCaseSensitivity = 1   ' 1 = Case Insensitivity, 0 = Case Sensitivity
  'uPrivilege = 1         ' 1 = ACL, 2 = Privileges, Else 0
  uZipFileName = FilePath & ZipName        ' The Zip File Name
  'uExtractDir            ' Extraction Directory, Null If Current Directory
  ' Create the file criteria.
  Crit = Array("@PSC_ReadMe*.txt")
  ' UnZip32.DLL will return the fist file found in the zip
  ' that matches the criteria.  Only one file is returned as
  ' the VBUnzBas UZDLLServ = 1 (abort)
  For I = 0 To UBound(Crit)
    If LenB(uZipInfo) = 0 Then ' Try next criteria   ' Try next criteria
      uZipNames.uzFiles(0) = Crit(I)
      uNumberFiles = 1
      VBUnZip32
    End If
  Next I
  ' To do, Control the size of uZipInfo.
  ' I intended this app to work as a 'quick view' for zips.  If the returned
  ' variable "uZipInfo" is on the large side it will slow the process.
  ' Need to limit the size of uZipInfo while the variable is built.
  If LenB(uZipInfo) Then
    txtSelectedReadme.Text = uZipInfo
    'actual code for getting the title
    strTitle = Left$(uZipInfo, InStr(1, uZipInfo, vbNewLine, vbTextCompare) - 1)
    strTitle = Replace$(strTitle, "Title: ", vbNullString, 1, 1, vbTextCompare)
    Else 'LENB(UZIPINFO) = FALSE/0
    lblFileName.Caption = "NO PSC README FOUND WITHIN THIS ARCHIVE!"
  End If
  'txtSelectedReadme.Visible = True
  'lblFileName.Visible = True
  Screen.MousePointer = vbDefault ' Return mouse pointer to normal.
  ProgName = strTitle
  txtProjectName.Text = strTitle

End Sub

':)Code Fixer V3.0.9 (7/19/2005 12:33:44 PM) 6 + 332 = 338 Lines Thanks Ulli for inspiration and lots of code.
