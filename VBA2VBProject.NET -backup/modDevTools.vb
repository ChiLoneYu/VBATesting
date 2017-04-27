Option Strict Off
Option Explicit On
Module modDevTools
	
	'Must have a comma at the end of these file extension lists for InStr to work correctly
	Public Const VBAImportTypes As String = "FRM,BAS,CLS,"
	Public Const VBAFileTypes As String = "FRMCMP,FRM,BAS,CLS,FRX,"
	Public Const VBAMoveTypes As String = "FRMCMP,BAS,CLS,VBP,"
	
	Dim project As Microsoft.Vbe.Interop.VBProject
	
	
	Public Sub ImportSourceFiles(ByRef sourcePath As String)
		Dim file As String
		Dim fileExt As String
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		file = Dir(sourcePath)
		While (file <> "")
			fileExt = GetFileExtension(file, True)
			'MsgBox fileExt
			If InStr(VBAImportTypes, fileExt & ",") > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.ActiveVBProject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AutoCADAcadApplication_definst.Application.VBE.ActiveVBProject.VBComponents.Import(sourcePath & file)
			End If
			
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			file = Dir()
		End While
	End Sub
	
	Public Sub ExportSourceFiles(ByRef destPath As String)
		Dim CRC32 As New clsCRC32
		Dim comp As Microsoft.Vbe.Interop.VBComponent
		Dim ignoreFilePath As String
		Dim cleanFilePath As String
		Dim ignorePath As String
		Dim fileSystemHandler As New Scripting.FileSystemObject
		Dim ignoreFile As Scripting.File
		Dim rootFile As Scripting.File
		Dim frxName As String
		Dim frmName As String
		Dim copyFRX As Boolean
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.ActiveVBProject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        project = AutoCADAcadApplication_definst.Application.VBE.ActiveVBProject

		
		ignorePath = destPath & "ignore\"
		'Clean out the ignore folder
		'fileSystemHandler.DeleteFile (ignorePath & "*")
		
		
		'No longer needed since modVfcMain is handling this...
		'Loop through all of the components and export to the ignore folder
		'    For Each comp In project.VBComponents
		'        If Left(comp.Name, 5) <> "Sheet" And comp.Name <> "ThisWorkbook" And comp.Name <> "DevTools" And comp.Name <> "ThisDrawing" Then
		'            'cleanFilePath = destPath & comp.Name & ToFileExtension(comp.Type)
		'            ignoreFilePath = ignorePath & comp.Name & ToFileExtension(comp.Type)
		'            comp.Export ignoreFilePath
		'        End If
		'    Next
		
		
		'Loop through root directory and find files with no match in ignore folder
		For	Each rootFile In fileSystemHandler.GetFolder(destPath).Files
			
			'Check that this file is a type handled by VBA
			If InStr(VBAFileTypes, GetFileExtension(rootFile.Name, True) & ",") > 0 Then
				
				ignoreFilePath = ignorePath & rootFile.Name
				
				'If file doesn't exist in the ignore folder, then go ahead and delete it... most likely removed from VBProject
				If Not fileSystemHandler.FileExists(ignoreFilePath) Then
					rootFile.Delete()
				End If
			End If
			
		Next rootFile
		
		'Loop through all of the files in the ignore directory and perform a CRC compare.  Copy to root if CRC doesn't match
		'Make sure to ignore .FRX
		
		For	Each ignoreFile In fileSystemHandler.GetFolder(ignorePath).Files
			
			copyFRX = False
			
			If InStr(VBAMoveTypes, GetFileExtension(ignoreFile.Name, True) & ",") > 0 Then
				'Check that file exists in root, otherwise it should be copied by default
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If Dir(destPath & ignoreFile.Name) <> "" Then
					'Move the file if the CRC32 doesn't match.. otherwise delete it
					If Hex(CRC32.CalcCRC32(ignorePath & ignoreFile.Name)) <> Hex(CRC32.CalcCRC32(destPath & ignoreFile.Name)) Then
						ignoreFile.Copy(destPath & ignoreFile.Name, True)
						
						'Also copy the matching FRM and FRX files if a form comparison file is found with a non-matching CRC32
						If GetFileExtension(ignoreFile.Name, True) = "FRMCMP" Then
							copyFRX = True
							
						End If
						
					End If
					
				Else
					'Copy from ignore to root since file doesn't exist
					ignoreFile.Copy(destPath, True)
					
					'Also copy the match FRX file if copying a form
					If GetFileExtension(ignoreFile.Name, True) = "FRMCMP" Then
						
						copyFRX = True
						
					End If
					
				End If
				
				'ignoreFile.Delete
				
			End If
			
			If copyFRX And GetFileExtension(ignoreFile.Name, True) = "FRMCMP" Then
				frxName = Left(ignoreFile.Name, Len(ignoreFile.Name) - 6) & "frx"
				frmName = Left(ignoreFile.Name, Len(ignoreFile.Name) - 6) & "frm"
				
				fileSystemHandler.CopyFile(ignorePath & frxName, destPath & frxName, True)
				fileSystemHandler.CopyFile(ignorePath & frmName, destPath & frmName, True)
				
				'fileSystemHandler.DeleteFile (ignorePath & frxName)
				'fileSystemHandler.DeleteFile (ignorePath & frmName)
				
			End If
			
			
		Next ignoreFile
		
		If frmVfcMain.cbDeleteIgnoreFiles.CheckState Then
			fileSystemHandler.DeleteFile(destPath & "ignore\*")
		End If
		
	End Sub
	
	Public Sub RemoveAllModules()
		Dim project As Microsoft.Vbe.Interop.VBProject
		'UPGRADE_WARNING: Couldn't resolve default property of object Application.VBE.ActiveVBProject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		project = AutoCADAcadApplication_definst.Application.VBE.ActiveVBProject
		
		MsgBox("Removing Modules")
		Dim comp As Microsoft.Vbe.Interop.VBComponent
		For	Each comp In project.VBComponents
			If Not comp.Name = "DevTools" And Left(comp.Name, 5) <> "Sheet" And comp.Name <> "ThisWorkbook" Then 'And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
				project.VBComponents.Remove(comp)
			End If
		Next comp
	End Sub
	
	Private Function ToFileExtension(ByRef vbeComponentType As Microsoft.Vbe.Interop.vbext_ComponentType) As String
		Select Case vbeComponentType
			Case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ClassModule
				ToFileExtension = ".cls"
			Case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule
				ToFileExtension = ".bas"
			Case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_MSForm
				ToFileExtension = ".frm"
			Case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_ActiveXDesigner
			Case Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_Document
			Case Else
				ToFileExtension = vbNullString
		End Select
		
	End Function
	
	Private Function GetFileExtension(ByRef fileName As String, Optional ByRef upperCase As Boolean = False) As String
		
		If upperCase Then
			GetFileExtension = UCase(Right(fileName, Len(fileName) - InStrRev(fileName, ".")))
		Else
			GetFileExtension = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
		End If
		
	End Function
End Module