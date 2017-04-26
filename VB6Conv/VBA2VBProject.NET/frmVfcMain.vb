Option Strict Off
Option Explicit On
Friend Class frmVfcMain
	Inherits System.Windows.Forms.Form
	
	
	Public ReadOnly Property SelectedProjectIndex() As Integer
		Get
			If lstVBProjects.SelectedIndex > -1 Then
				SelectedProjectIndex = CInt(lstVBProjects.Text)
			End If
		End Get
	End Property
	
	Public ReadOnly Property SelectedProjectFilename() As String
		Get
			If lstVBProjects.SelectedIndex > -1 Then
				SelectedProjectFilename = lstVBProjects.List(lstVBProjects.SelectedIndex, 2)
			End If
		End Get
	End Property
	
	Public ReadOnly Property SelectedProjectGitDirectory() As String
		Get
			If lstVBProjects.SelectedIndex > -1 Then
                SelectedProjectGitDirectory = txtGitRepoPath.Text & StripExtensionFromFileName(GetFileNameFromPath(lstVBProjects.SelectedItem(lstVBProjects.SelectedIndex, 2)))
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If Dir(SelectedProjectGitDirectory, FileAttribute.Directory) = "" Then MkDir(SelectedProjectGitDirectory)
			End If
		End Get
	End Property
	
	Public ReadOnly Property SelectedProjectGitIgnoreDirectory() As String
		Get
			If lstVBProjects.SelectedIndex > -1 Then
				SelectedProjectGitIgnoreDirectory = SelectedProjectGitDirectory & "\ignore\"
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If Dir(SelectedProjectGitIgnoreDirectory, FileAttribute.Directory) = "" Then MkDir(SelectedProjectGitIgnoreDirectory)
			End If
		End Get
	End Property
	
	
	Public ReadOnly Property IncludeCode() As Boolean
		Get
			IncludeCode = (chkIncludeCode.CheckState = True)
		End Get
	End Property
	
	Public ReadOnly Property ShowUnknown() As Boolean
		Get
			ShowUnknown = (chkShowUnknown.CheckState = True)
		End Get
	End Property
	
	Private Sub cmdBrowseRepo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseRepo.Click
		
		txtGitRepoPath.Text = BrowseFolder("Select GIT Repo Base Folder", Environ("USERPROFILE") & "\Source\Repos\Autocad Automation\")
		'MsgBox Environ("USERPROFILE")
		'"C:\Users\acunningham\Source\Repos"
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdConvert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConvert.Click
		cmdCancel.Text = "Exit"
		Call ProcessProject()
		Call ExportSourceFiles(SelectedProjectGitDirectory & "\")
	End Sub
	
	
	
	'UPGRADE_WARNING: Event lstVBProjects.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstVBProjects_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstVBProjects.SelectedIndexChanged
		Dim objProj As Microsoft.Vbe.Interop.VBProject
		Dim intIndex As Short
		
		If lstVBProjects.SelectedIndex > -1 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.VBProjects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			objProj = AutoCADAcadApplication_definst.Application.VBE.VBProjects(Me.SelectedProjectIndex)
			cmdConvert.Enabled = True
			
		End If
		
		'UPGRADE_NOTE: Object objProj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objProj = Nothing
	End Sub
	
	
	
	Private Sub UserForm_Activate()
		Dim objProjs As Microsoft.Vbe.Interop.VBProjects
		Dim objProj As Microsoft.Vbe.Interop.VBProject
		Dim intIndex As Short
		Dim strTemp As String
		
		On Error GoTo ErrorHandler
		
		'Load up list box with VBProjects
		'UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.VBProjects. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		objProjs = AutoCADAcadApplication_definst.Application.VBE.VBProjects
		cmdConvert.Enabled = False
		'  lstMSForms.Clear
		lstVBProjects.Items.Clear()
		'UPGRADE_WARNING: Couldn't resolve default property of object lstVBProjects.ColumnCount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lstVBProjects.ColumnCount = 3
		'UPGRADE_WARNING: Couldn't resolve default property of object lstVBProjects.ColumnWidths. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lstVBProjects.ColumnWidths = "20 pt;90 pt;400 pt"
		Dim listIndex As Short
		listIndex = 1
		For intIndex = 1 To objProjs.Count
			objProj = objProjs.Item(intIndex)
			If (objProj.fileName <> objProjs.VBE.ActiveVBProject.fileName) Then 'do not list the exporter macro
				lstVBProjects.Items.Add(CStr(intIndex))
				lstVBProjects.List(listIndex - 1, 1) = objProj.Name
				lstVBProjects.List(listIndex - 1, 2) = objProj.fileName
				listIndex = listIndex + 1
			End If
		Next intIndex
		txtGitRepoPath.Text = Environ("USERPROFILE") & "\Source\Repos\Autocad Automation\"
		
		
SubExit: 
		'UPGRADE_NOTE: Object objProj may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objProj = Nothing
		'UPGRADE_NOTE: Object objProjs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objProjs = Nothing
		Exit Sub
ErrorHandler: 
		Select Case Err.Number
			Case 76 'Path not found
				strTemp = "VFC can only work with SAVED projects," & vbCrLf & "please save all newly created projects" & vbCrLf & "and start again."
				Call MsgBox(strTemp, MsgBoxStyle.Critical, ccAPPNAME)
			Case Else 'Huh??
				strTemp = ccAPPNAME & " ERROR" & vbCrLf & "VFC_001: Error in UserForm_Activate" & vbCrLf & "Description: " & Err.Description & vbCrLf & "Source: " & Err.Source & vbCrLf & "Number: " & CStr(Err.Number) & vbCrLf & "VFC Ver: " & ccAPPVER
				Call MsgBox(strTemp, MsgBoxStyle.Critical, ccAPPNAME)
		End Select
		Me.Close()
		Resume SubExit
	End Sub
	
	Private Sub UserForm_Terminate()
		Debug.Print("UserForm_Terminate frmVfcMain")
	End Sub
End Class