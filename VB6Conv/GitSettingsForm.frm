VERSION 5.00
Begin VB.Form GitSettingsForm
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7320
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   8580
   BeginProperty Font
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0  'False
      Italic          =   0  'False
      Strikethrough   =   0  'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1
      Caption         =   "ShibbyGit Settings"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   0  'False
         Strikethrough   =   0  'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2760
   End
   Begin VB.CommandButton CancelButton
      Caption         =   "Cancel"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   0  'False
         Strikethrough   =   0  'False
      EndProperty
      Height          =   480
      Left            =   5160
      TabIndex        =   1
      Top             =   6600
      Width           =   1560
   End
   Begin VB.CommandButton OKButton
      Caption         =   "OK"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   0  'False
         Strikethrough   =   0  'False
      EndProperty
      Height          =   480
      Left            =   1560
      TabIndex        =   2
      Top             =   6600
      Width           =   1560
   End
   Begin VB.Frame ProjectSettingsFrame
      Caption         =   "Project Settings"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   -1  'True
         Strikethrough   =   0  'False
      EndProperty
      Height          =   2520
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   8280
      Begin VB.Label Label3
         Caption         =   "Git project path"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   270
         TabIndex        =   0
         Top             =   480
         Width           =   2160
      End
      Begin VB.TextBox ProjectPathTextBox
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   2550
         TabIndex        =   1
         Top             =   480
         Width           =   4080
      End
      Begin VB.CommandButton ProjectPathBrowseButton
         Caption         =   "Browse"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   480
         Left            =   6750
         TabIndex        =   2
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label4
         Caption         =   "Project Structure"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   270
         TabIndex        =   3
         Top             =   960
         Width           =   2040
      End
      Begin VB.Label Label5
         Caption         =   "User Name"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   270
         TabIndex        =   4
         Top             =   1440
         Width           =   2040
      End
      Begin VB.TextBox UserNameBox
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   2550
         TabIndex        =   5
         Top             =   1440
         Width           =   4080
      End
      Begin VB.Label Label6
         Caption         =   "User Email"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   270
         TabIndex        =   6
         Top             =   1920
         Width           =   2040
      End
      Begin VB.TextBox UserEmailBox
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   2550
         TabIndex        =   7
         Top             =   1920
         Width           =   4080
      End
      Begin VB.ComboBox FileStructureBox
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   2880
      End
   End
   Begin VB.Frame IOSettingsFrame
      Caption         =   "IO Settings"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   -1  'True
         Strikethrough   =   0  'False
      EndProperty
      Height          =   1440
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   8280
      Begin VB.CheckBox ExportOnGitBox
         Caption         =   "Auto-Export Source Before Status or Console?"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   420
         Left            =   510
         TabIndex        =   0
         Top             =   360
         Width           =   5820
      End
      Begin VB.CheckBox FrxCleanupBox
         Caption         =   "Clean up unnessecary .frx files on export?"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   405
         Left            =   510
         TabIndex        =   1
         Top             =   849.000015258789
         Width           =   5820
      End
   End
   Begin VB.Frame GlobalSettingsFrame
      Caption         =   "Global Settings"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   -1  'True
         Strikethrough   =   0  'False
      EndProperty
      Height          =   1200
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   8280
      Begin VB.Label Label2
         Caption         =   "Git executable path"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   270
         TabIndex        =   0
         Top             =   480
         Width           =   2160
      End
      Begin VB.TextBox GitExeTextBox
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   360
         Left            =   2550
         TabIndex        =   1
         Top             =   480
         Width           =   4080
      End
      Begin VB.CommandButton GitExeBrowseButton
         Caption         =   "Browse"
         BeginProperty Font
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0  'False
            Italic          =   0  'False
            Strikethrough   =   0  'False
         EndProperty
         Height          =   480
         Left            =   6750
         TabIndex        =   2
         Top             =   360
         Width           =   1440
      End
   End
End
Attribute VB_Name = "GitSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private needGitUserNameUpdate As Boolean
Private needGitUserEmailUpdate As Boolean


'****************************************************************
' initialize

Public Sub resetForm()
    ' set the gitExe path text
    GitExeTextBox.Text = ShibbySettings.GitExePath
    
    ' set the project path text
    ProjectPathTextBox.Text = ShibbySettings.GitProjectPath
    
    If GitExeTextBox.Text <> "" Then
        ' set the username field
        Dim userName As String
        If ProjectPathTextBox.Text = "" Then
            userName = GitCommands.RunGitAsProcess("config user.name", UseProjectPath:=False)
        Else
            userName = GitCommands.RunGitAsProcess("config user.name")
        End If
        If Len(userName) > 0 Then
            userName = Left(userName, Len(userName) - 1)
        End If
        UserNameBox.value = userName
        
        ' set the email field
        Dim userEmail As String
        If ProjectPathTextBox.Text = "" Then
            userEmail = GitCommands.RunGitAsProcess("config user.email", UseProjectPath:=False)
        Else
            userEmail = GitCommands.RunGitAsProcess("config user.email")
        End If
        If Len(userEmail) > 0 Then
            userEmail = Left(userEmail, Len(userEmail) - 1)
        End If
        UserEmailBox.value = userEmail
    End If
    
    ' set the frx box value
    FrxCleanupBox.value = ShibbySettings.FrxCleanup
    
    ' set the frx box value
    ExportOnGitBox.value = ShibbySettings.ExportOnGit
    
    ' Add items to the file structure box
    FileStructureBox.AddItem "Flat File Stucture"
    FileStructureBox.AddItem "Simple Src Structure"
    FileStructureBox.AddItem "Separated Src Structure"
    Dim fsIndex As ShibbyFileStructure
    fsIndex = ShibbySettings.fileStructure
    FileStructureBox.ListIndex = fsIndex
    
    needGitUserNameUpdate = False
    needGitUserEmailUpdate = False
    
End Sub


'****************************************************************
' component callbacks

Private Sub CancelButton_Click()
    GitSettingsForm.hide
End Sub

Private Sub OKButton_Click()
    SaveGitExe
    SaveProjectPath
    SaveUserName
    SaveUserEmail
    SaveFrxCleanup
    SaveExportOnGit
    SaveFileStructure
    GitSettingsForm.hide
End Sub

Private Sub UserEmailBox_Change()
    needGitUserEmailUpdate = True
End Sub

Private Sub UserNameBox_Change()
    needGitUserNameUpdate = True
End Sub


Private Sub GitExeBrowseButton_Click()
    GitExeTextBox.Text = FileUtils.FileBrowser("Browser for git.exe")
End Sub


Private Sub ProjectPathBrowseButton_Click()
    ProjectPathTextBox.Text = FileUtils.FolderBrowser("Browse for Git project folder")
End Sub


'****************************************************************
' save methods

' Save the project path as a document property
Private Sub SaveProjectPath()
    Dim newPath As String
    newPath = ProjectPathTextBox.Text
    
    If newPath <> "" And FileUtils.FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find file: " & newPath
        Exit Sub
    End If

    'save this one in the registry
    ShibbySettings.GitProjectPath = newPath
End Sub


' save the gitExe path as a registry property
Private Sub SaveGitExe()
    Dim newPath As String
    newPath = GitExeTextBox.Text
    
    If newPath <> "" And FileUtils.FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find file: " & newPath
        Exit Sub
    End If

    'save this one in the registry
    ShibbySettings.GitExePath = newPath
End Sub

' save the user email to the git repo
Private Sub SaveUserEmail()
    If needGitUserEmailUpdate Then
        GitCommands.RunGitAsProcess ("config --local user.email """ & UserEmailBox.value & """")
    End If
    needGitUserEmailUpdate = False
End Sub


' save the user name to the git repo
Private Sub SaveUserName()
    If needGitUserNameUpdate Then
        GitCommands.RunGitAsProcess ("config --local user.name """ & UserNameBox.value & """")
    End If
    needGitUserNameUpdate = False
End Sub

' save the frx setting
Private Sub SaveFrxCleanup()
    ShibbySettings.FrxCleanup = FrxCleanupBox.value
End Sub

' save the export on git setting
Private Sub SaveExportOnGit()
    ShibbySettings.ExportOnGit = ExportOnGitBox.value
End Sub

' save the File structure
Private Sub SaveFileStructure()
    ShibbySettings.fileStructure = FileStructureBox.ListIndex
End Sub
