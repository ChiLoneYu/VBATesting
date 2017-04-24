VERSION 5.00
Begin VB.Form GitConsoleForm
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6084
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   6204
   BeginProperty Font
      Name            =   "Consolas"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0  'False
      Italic          =   0  'False
      Strikethrough   =   0  'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6084
   ScaleWidth      =   6204
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox CommandBox
      BackColor       =   &H0&
      ForeColor       =   &HFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5640
   End
   Begin VB.Label TitleLabel
      Alignment       =   2  'Center
      Caption         =   "ShibbyGit VB Git Console"
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
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   6240
   End
   Begin VB.TextBox OutputBox
      BackColor       =   &H0&
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   0  'False
         Strikethrough   =   0  'False
      EndProperty
      ForeColor       =   &HFF00&
      Height          =   4575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "ShibbyGit Console                                                                                                                              .                                                                                                                              Enter git commands and press enter, with or without ""git""                                                       e.g. ""log"", ""git status"", or ""branch""                                                                                                                                                                           .                                                                                                                                                                      .                                                                                                                                                                       Hold the Shift key to run command in a cmd shell                                                                                    .                                                                                                                              .                                                                                                                                                                        Must run in cmd shell for web operations, e.g. GitHub remote access"
      Top             =   1320
      Width           =   5640
   End
End
Attribute VB_Name = "GitConsoleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit
Private CommandHistory As New Collection
Private CommandIndex As Integer

' execute command when enter is pressed
Private Sub CommandBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    ' commandIndex checking
    If CommandIndex <= 0 Then
        CommandIndex = 1
    End If
    
    If CommandIndex > CommandHistory.Count Then
        CommandIndex = CommandHistory.Count
    End If
    
    ' Add a blank item if empty commandHistory
    If CommandHistory.Count = 0 Then
        CommandHistory.Add ""
    End If

    ' return key: process command
    If KeyCode = vbKeyReturn Then
     
        Dim useShell As Boolean
        useShell = (Shift = 1)
             
        ' allow "git " to preceed options, for muscle memory!
        ' process "export" and "import" differently
        Dim output As String
        If CommandBox.Text Like "git *" Then
            CommandBox.Text = Right(CommandBox.Text, Len(CommandBox.Text) - 4)
        End If
        
        ' parse for available options
        If CommandBox.Text = "export" Then
            output = GitIO.GitExport(ShibbySettings.GitProjectPath, ShibbySettings.fileStructure)
        ElseIf CommandBox.Text = "import" Then
            output = GitIO.GitImport(ShibbySettings.GitProjectPath, ShibbySettings.fileStructure)
        Else
            If useShell Then
                output = "Shell exectution"
                GitCommands.RunGitInShell (CommandBox.Text)
            Else
                output = GitCommands.RunGitAsProcess(CommandBox.Text, 1500)
            End If
        End If
        
        ' push the command on the history
        If CommandBox.Text <> CommandHistory.Item(CommandIndex) Then
            CommandHistory.Add CommandBox.Text, After:=CommandIndex
            CommandIndex = CommandIndex + 1
        End If
        
        ' display the output
        OutputBox.value = output
        OutputBox.SelLength = 0
        OutputBox.SelStart = 0
        OutputBox.SetFocus
        KeyCode.value = 0
        
    ' up key: show previous command
    ElseIf KeyCode = vbKeyUp Then
        If CommandIndex > 1 Then
            CommandIndex = CommandIndex - 1
        End If
        CommandBox.Text = CommandHistory(CommandIndex)
        KeyCode.value = 0
        
    ' down key: show next command
    ElseIf KeyCode = vbKeyDown Then
        If CommandIndex < CommandHistory.Count Then
            CommandIndex = CommandIndex + 1
        End If
        CommandBox.Text = CommandHistory(CommandIndex)
        KeyCode.value = 0
  End If

End Sub


Private Sub CommandBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GiveCommandBoxFocusAndSelect
    End If
End Sub

Private Sub OutputBox_AfterUpdate()
    GiveCommandBoxFocusAndSelect
End Sub


Private Sub OutputBox_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Cancel = True
    GiveCommandBoxFocusAndSelect
End Sub

Private Sub OutputBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    GiveCommandBoxFocusAndSelect
End Sub


Private Sub GiveCommandBoxFocusAndSelect()
    CommandBox.SetFocus
    CommandBox.SelStart = 0
    CommandBox.SelLength = Len(CommandBox.value)
End Sub

Private Sub OutputBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button = vbKeyMButton Then
        With CommandBox
            .SelText = OutputBox.SelText
            .SetFocus
        End With
    End If
End Sub

Private Sub CommandBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button = vbKeyMButton Then
        CommandBox.SelText = OutputBox.SelText
    End If
End Sub

