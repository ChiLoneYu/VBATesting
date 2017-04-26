Option Strict Off
Option Explicit On
Module modFileFunctions
	Private Const BIF_RETURNONLYFSDIRS As Integer = &H1
	Private Const BIF_DONTGOBELOWDOMAIN As Integer = &H2
	Private Const BIF_RETURNFSANCESTORS As Integer = &H8
	Private Const BIF_BROWSEFORCOMPUTER As Integer = &H1000
	Private Const BIF_BROWSEFORPRINTER As Integer = &H2000
	Private Const BIF_BROWSEINCLUDEFILES As Integer = &H4000
	Private Const MAX_PATH As Integer = 260
	
	Function BrowseFolder(Optional ByRef Caption As String = "", Optional ByRef InitialFolder As String = "") As String
		
		Dim SH As Shell32.Shell
		Dim F As Shell32.Folder
		
		SH = New Shell32.Shell
		F = SH.BrowseForFolder(0, Caption, BIF_RETURNONLYFSDIRS, InitialFolder)
		If Not F Is Nothing Then
            Return F.Items.Item.Path
		End If
		
	End Function
	
	'We store our output files in a subdirectory of the original project folder
	'  so added a subfolder to the returned path.
	Public Function GetFileDirectory(ByRef strFileName As String) As String
        Dim tmpString As String
        Dim intPos As Short
		For intPos = Len(strFileName) To 1 Step -1
			If Mid(strFileName, intPos, 1) = "\" Then
                tmpString = Left(strFileName, intPos)
				Exit For
			End If
		Next intPos
        tmpString = tmpString & "VB6Conv\"
		'Dim fso As New FileSystemObject
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        If Dir(tmpString, FileAttribute.Directory) = "" Then MkDir(GetFileDirectory)
        Return tmpString

	End Function
	
	Public Function GetFileNameFromPath(ByVal strFilePath As String) As String
		GetFileNameFromPath = Mid(strFilePath, InStrRev(strFilePath, "\") + 1)
	End Function
	
	Public Function StripExtensionFromFileName(ByVal strFileName As String) As String
		StripExtensionFromFileName = Left(strFileName, InStrRev(strFileName, ".") - 1)
		
	End Function
End Module