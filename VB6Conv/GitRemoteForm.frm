VERSION 5.00
Begin VB.Form GitRemoteForm
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   5415
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
   ScaleHeight     =   3660
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1
      Alignment       =   2  'Center
      Caption         =   "Remote Options"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   15.75
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
      Width           =   5160
   End
   Begin VB.ListBox RemoteBox
      Height          =   960
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   2280
   End
   Begin VB.ListBox BranchBox
      Height          =   960
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   2280
   End
   Begin VB.Label Label2
      Caption         =   "Remote"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label3
      Caption         =   "Branch"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   1080
   End
   Begin VB.CheckBox PushButton
      Caption         =   "Push"
      Height          =   405
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Value           =   1  'Checked
      Width           =   765
   End
   Begin VB.CheckBox PullButton
      Caption         =   "Pull"
      Height          =   405
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   750
   End
   Begin VB.CheckBox FetchButton
      Caption         =   "Fetch"
      Height          =   405
      Left            =   3120
      TabIndex        =   7
      Top             =   840
      Width           =   810
   End
   Begin VB.CommandButton OKButton
      Caption         =   "Run"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   0  'False
         Strikethrough   =   0  'False
      EndProperty
      Height          =   600
      Left            =   360
      TabIndex        =   8
      Top             =   2880
      Width           =   2160
   End
   Begin VB.CommandButton DoneButton
      Caption         =   "Done"
      BeginProperty Font
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0  'False
         Italic          =   0  'False
         Strikethrough   =   0  'False
      EndProperty
      Height          =   600
      Left            =   3000
      TabIndex        =   9
      Top             =   2880
      Width           =   2160
   End
End
Attribute VB_Name = "GitRemoteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False















Public remotes As Collection
Public branches As Collection

Public Sub resetForm()
    Set branches = GitParser.ParseBranches
    AddBranchesToList
    
    Set remotes = GitParser.ParseRemotes
    AddPushRemotesToList
End Sub


Private Sub AddPushRemotesToList()
    Dim currentInd As Integer
    currentInd = RemoteBox.ListIndex
    
    RemoteBox.Clear
    Dim remote As GitRemote
    For Each remote In remotes
        If remote.RemoteType = "push" Then
            RemoteBox.AddItem remote.name
        End If
    Next remote
    
    If currentInd <= UBound(RemoteBox.List) And currentInd > LBound(RemoteBox.List) Then
        RemoteBox.ListIndex = currentInd
    ElseIf RemoteBox.ListCount > 0 Then
        RemoteBox.ListIndex = 0
    End If
End Sub

Private Sub AddFetchRemotesToList()
    Dim currentInd As Integer
    currentInd = RemoteBox.ListIndex
    
    RemoteBox.Clear
    Dim remote As GitRemote
    For Each remote In remotes
        If remote.RemoteType = "fetch" Then
            RemoteBox.AddItem remote.name
        End If
    Next remote
    
    If currentInd <= UBound(RemoteBox.List) And currentInd > LBound(RemoteBox.List) Then
        RemoteBox.ListIndex = currentInd
    ElseIf RemoteBox.ListCount > 0 Then
        RemoteBox.ListIndex = 0
    End If
End Sub

Private Sub AddBranchesToList()
    Dim currentInd As Integer
    currentInd = BranchBox.ListIndex

    BranchBox.Clear
    Dim br As GitBranch
    For Each br In branches
        If br.Active Then
            BranchBox.AddItem "*" & br.name, 0
        Else
            BranchBox.AddItem br.name
        End If
    Next br
    
    If currentInd <= UBound(BranchBox.List) And currentInd > LBound(BranchBox.List) Then
        BranchBox.ListIndex = currentInd
    Else
        BranchBox.ListIndex = 0
    End If
End Sub


Private Sub DoneButton_Click()
    GitRemoteForm.hide
End Sub

Private Sub OKButton_Click()
    
    If RemoteBox.ListIndex = -1 Or BranchBox.ListIndex = -1 Then
        Exit Sub
    End If
    
    Dim operation As String
    If PushButton.value = True Then
        operation = "push"
    ElseIf PullButton.value = True Then
        operation = "pull"
    Else
        operation = "fetch"
    End If
    
    Dim remote As String
    remote = RemoteBox.List(RemoteBox.ListIndex)
    
    Dim branch As String
    branch = BranchBox.List(BranchBox.ListIndex)
    branch = Replace(branch, "*", "")
    
    Dim gitParms As String
    gitParms = operation & " " & remote & " " & branch
    
    GitCommands.RunGitInShell gitParms
End Sub

Private Sub PushButton_Click()
    AddPushRemotesToList
End Sub

Private Sub PullButton_Click()
    AddFetchRemotesToList
End Sub

Private Sub FetchButton_Click()
    AddFetchRemotesToList
End Sub
