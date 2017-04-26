VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVfcMain 
   Caption         =   "VBA Source Control Export"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   OleObjectBlob   =   "frmVfcMain.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "VFC"
End
Attribute VB_Name = "frmVfcMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Property Get SelectedProjectIndex() As Long
  If lstVBProjects.listIndex > -1 Then
    SelectedProjectIndex = CLng(lstVBProjects.Text)
  End If
End Property

Public Property Get SelectedProjectFilename() As String
  If lstVBProjects.listIndex > -1 Then
    SelectedProjectFilename = lstVBProjects.List(lstVBProjects.listIndex, 2)
  End If
End Property

Public Property Get SelectedProjectGitDirectory() As String
  If lstVBProjects.listIndex > -1 Then
    SelectedProjectGitDirectory = txtGitRepoPath & StripExtensionFromFileName(GetFileNameFromPath(lstVBProjects.List(lstVBProjects.listIndex, 2)))
    If Dir(SelectedProjectGitDirectory, vbDirectory) = "" Then MkDir SelectedProjectGitDirectory
  End If
End Property

Public Property Get SelectedProjectGitIgnoreDirectory() As String
  If lstVBProjects.listIndex > -1 Then
    SelectedProjectGitIgnoreDirectory = SelectedProjectGitDirectory & "\ignore\"
    If Dir(SelectedProjectGitIgnoreDirectory, vbDirectory) = "" Then MkDir SelectedProjectGitIgnoreDirectory
  End If
End Property


Public Property Get IncludeCode() As Boolean
  IncludeCode = (chkIncludeCode.Value = True)
End Property

Public Property Get ShowUnknown() As Boolean
  ShowUnknown = (chkShowUnknown.Value = True)
End Property

Private Sub cmdBrowseRepo_Click()

    txtGitRepoPath = BrowseFolder("Select GIT Repo Base Folder", Environ("USERPROFILE") & "\Source\Repos\Autocad Automation\")
    'MsgBox Environ("USERPROFILE")
    '"C:\Users\acunningham\Source\Repos"
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdConvert_Click()
  cmdCancel.Caption = "Exit"
  Call ProcessProject
  Call ExportSourceFiles(SelectedProjectGitDirectory & "\")
End Sub



Private Sub lstVBProjects_Click()
  Dim objProj As VBIDE.VBProject
  Dim intIndex As Integer
  
  If lstVBProjects.listIndex > -1 Then
    Set objProj = Application.VBE.VBProjects(Me.SelectedProjectIndex)
    cmdConvert.Enabled = True

  End If
  
  Set objProj = Nothing
End Sub



Private Sub UserForm_Activate()
  Dim objProjs As VBProjects
  Dim objProj As VBProject
  Dim intIndex As Integer
  Dim strTemp As String
  
  On Error GoTo ErrorHandler
  
  'Load up list box with VBProjects
  Set objProjs = Application.VBE.VBProjects
  cmdConvert.Enabled = False
'  lstMSForms.Clear
  lstVBProjects.Clear
  lstVBProjects.ColumnCount = 3
  lstVBProjects.ColumnWidths = "20 pt;90 pt;400 pt"
  Dim listIndex As Integer
  listIndex = 1
  For intIndex = 1 To objProjs.Count
    Set objProj = objProjs(intIndex)
    If (objProj.Name <> "ExportToVB6") Then 'do not list the exporter macro
        lstVBProjects.AddItem CStr(intIndex)
        lstVBProjects.List(listIndex - 1, 1) = objProj.Name
        lstVBProjects.List(listIndex - 1, 2) = objProj.fileName
        listIndex = listIndex + 1
    End If
  Next intIndex
  txtGitRepoPath = Environ("USERPROFILE") & "\Source\Repos\Autocad Automation\"
  
  
SubExit:
  Set objProj = Nothing
  Set objProjs = Nothing
  Exit Sub
ErrorHandler:
  Select Case Err.Number
    Case 76 'Path not found
      strTemp = "VFC can only work with SAVED projects," & vbCrLf & _
                "please save all newly created projects" & vbCrLf & _
                "and start again."
      Call MsgBox(strTemp, vbCritical, ccAPPNAME)
    Case Else 'Huh??
      strTemp = ccAPPNAME & " ERROR" & vbCrLf & _
                "VFC_001: Error in UserForm_Activate" & vbCrLf & _
                "Description: " & Err.Description & vbCrLf & _
                "Source: " & Err.Source & vbCrLf & _
                "Number: " & CStr(Err.Number) & vbCrLf & _
                "VFC Ver: " & ccAPPVER
      Call MsgBox(strTemp, vbCritical, ccAPPNAME)
  End Select
  Unload Me
  Resume SubExit
End Sub

Private Sub UserForm_Terminate()
  Debug.Print "UserForm_Terminate frmVfcMain"
End Sub


