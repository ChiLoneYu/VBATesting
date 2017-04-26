Attribute VB_Name = "modFileFunctions"
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN As Long = &H2
Private Const BIF_RETURNFSANCESTORS As Long = &H8
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000
Private Const BIF_BROWSEFORPRINTER As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000
Private Const MAX_PATH As Long = 260

Function BrowseFolder(Optional Caption As String, _
    Optional InitialFolder As String) As String

Dim SH As Shell32.Shell
Dim F As Shell32.Folder

Set SH = New Shell32.Shell
Set F = SH.BrowseForFolder(0&, Caption, BIF_RETURNONLYFSDIRS, InitialFolder)
If Not F Is Nothing Then
    BrowseFolder = F.Items.Item.Path
End If

End Function

'We store our output files in a subdirectory of the original project folder
'  so added a subfolder to the returned path.
Public Function GetFileDirectory(strFileName As String) As String
  Dim intPos As Integer
  For intPos = Len(strFileName) To 1 Step -1
    If Mid(strFileName, intPos, 1) = "\" Then
      GetFileDirectory = Left(strFileName, intPos)
      Exit For
    End If
  Next intPos
  GetFileDirectory = GetFileDirectory & "VB6Conv\"
  'Dim fso As New FileSystemObject
  If Dir(GetFileDirectory, vbDirectory) = "" Then MkDir GetFileDirectory
End Function

Public Function GetFileNameFromPath(ByVal strFilePath As String) As String
    GetFileNameFromPath = Mid(strFilePath, InStrRev(strFilePath, "\") + 1)
End Function

Public Function StripExtensionFromFileName(ByVal strFileName As String) As String
    StripExtensionFromFileName = Left(strFileName, InStrRev(strFileName, ".") - 1)

End Function
