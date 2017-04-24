VERSION 5.00
Begin VB.Form GitCommitMessageForm
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2256
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   7512
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
   ScaleHeight     =   2256
   ScaleWidth      =   7512
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label TitleLabel
      Caption         =   "Enter Commit Message"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   3360
   End
   Begin VB.TextBox MessageTextBox
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   0  'False
         Strikethrough   =   0  'False
      EndProperty
      Height          =   348
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   6720
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
      Left            =   4800
      TabIndex        =   2
      Top             =   1320
      Width           =   1320
   End
   Begin VB.CommandButton OKButton
      Caption         =   "commit -am"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1440
   End
End
Attribute VB_Name = "GitCommitMessageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













Private Sub CancelButton_Click()
    GitCommitMessageForm.hide
End Sub


Private Sub MessageTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OKButton.SetFocus
        DoEvents
        OKButton_Click
    End If
End Sub

Private Sub OKButton_Click()
    Dim commitMessage As String
    commitMessage = MessageTextBox.Text
    
    If commitMessage = "" Then
        MsgBox "Please enter a commit message"
        Exit Sub
    End If
    
    GitCommands.GitCommit (commitMessage)
    GitCommitMessageForm.hide
End Sub

