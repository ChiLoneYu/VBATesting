Attribute VB_Name = "modDevTools"
Option Explicit

'Must have a comma at the end of these file extension lists for InStr to work correctly
Public Const VBAImportTypes = "FRM,BAS,CLS,"
Public Const VBAFileTypes = "FRMCMP,FRM,BAS,CLS,FRX,"
Public Const VBAMoveTypes = "FRMCMP,BAS,CLS,VBP,"

Dim project As VBProject


Public Sub ImportSourceFiles(sourcePath As String)
    Dim file As String
    Dim fileExt As String
    file = Dir(sourcePath)
    While (file <> "")
        fileExt = GetFileExtension(file, True)
        'MsgBox fileExt
        If InStr(VBAImportTypes, fileExt & ",") > 0 Then
            Application.VBE.ActiveVBProject.VBComponents.Import sourcePath & file
        End If
        
        file = Dir
    Wend
End Sub

Public Sub ExportSourceFiles(destPath As String)
    Dim CRC32 As New clsCRC32
    Dim comp As VBComponent
    Dim ignoreFilePath As String
    Dim cleanFilePath As String
    Dim ignorePath As String
    Dim fileSystemHandler As New Scripting.FileSystemObject
    Dim ignoreFile As Scripting.file
    Dim rootFile As Scripting.file
    Dim frxName As String
    Dim frmName As String
    Dim copyFRX As Boolean
    
    
    Set project = Application.VBE.ActiveVBProject
    
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
    For Each rootFile In fileSystemHandler.GetFolder(destPath).Files
        
        'Check that this file is a type handled by VBA
        If InStr(VBAFileTypes, GetFileExtension(rootFile.Name, True) & ",") > 0 Then
        
            ignoreFilePath = ignorePath & rootFile.Name
            
            'If file doesn't exist in the ignore folder, then go ahead and delete it... most likely removed from VBProject
            If Not fileSystemHandler.FileExists(ignoreFilePath) Then
                rootFile.Delete
            End If
        End If
        
    Next
    
    'Loop through all of the files in the ignore directory and perform a CRC compare.  Copy to root if CRC doesn't match
    'Make sure to ignore .FRX
    
    For Each ignoreFile In fileSystemHandler.GetFolder(ignorePath).Files
        
        copyFRX = False
        
        If InStr(VBAMoveTypes, GetFileExtension(ignoreFile.Name, True) & ",") > 0 Then
            'Check that file exists in root, otherwise it should be copied by default
            If Dir(destPath & ignoreFile.Name) <> "" Then
                'Move the file if the CRC32 doesn't match.. otherwise delete it
                If Hex$(CRC32.CalcCRC32(ignorePath & ignoreFile.Name)) <> Hex$(CRC32.CalcCRC32(destPath & ignoreFile.Name)) Then
                    ignoreFile.Copy destPath & ignoreFile.Name, True
                    
                    'Also copy the matching FRM and FRX files if copying a form comparison file
                    If GetFileExtension(ignoreFile.Name, True) = "FRMCMP" Then
                        copyFRX = True
                        
                    End If

                End If
                
                
            Else
                'Copy from ignore to root since file doesn't exist
                ignoreFile.Copy destPath, True
                
            
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
                
                fileSystemHandler.CopyFile ignorePath & frxName, destPath & frxName, True
                fileSystemHandler.CopyFile ignorePath & frmName, destPath & frmName, True
                
                'fileSystemHandler.DeleteFile (ignorePath & frxName)
                'fileSystemHandler.DeleteFile (ignorePath & frmName)
                
        End If
        
        
    Next
    
   
End Sub

Public Sub RemoveAllModules()
    Dim project As VBProject
    Set project = Application.VBE.ActiveVBProject
    
    MsgBox "Removing Modules"
    Dim comp As VBComponent
    For Each comp In project.VBComponents
        If Not comp.Name = "DevTools" And Left(comp.Name, 5) <> "Sheet" And comp.Name <> "ThisWorkbook" Then 'And (comp.Type = vbext_ct_ClassModule Or comp.Type = vbext_ct_StdModule) Then
            project.VBComponents.Remove comp
        End If
    Next
End Sub

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select
    
End Function

Private Function GetFileExtension(fileName As String, Optional upperCase As Boolean = False) As String
    
    If upperCase Then
        GetFileExtension = UCase(Right(fileName, Len(fileName) - InStrRev(fileName, ".")))
    Else
        GetFileExtension = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
    End If

End Function

