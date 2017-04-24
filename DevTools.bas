Attribute VB_Name = "DevTools"
Option Explicit

Public Const IgnoreImport = "frx,log"
Public Const GoodImport = "FRM,BAS,CLS"
Public Const VBAFileTypes = "FRM,BAS,CLS,FRX"

Dim project As VBProject


Public Sub ImportSourceFiles(sourcePath As String)
    Dim file As String
    Dim fileExt As String
    file = Dir(sourcePath)
    While (file <> "")
        fileExt = GetFileExtension(file, True)
        'MsgBox fileExt
        If InStr(GoodImport, fileExt) > 0 Then
            Application.VBE.ActiveVBProject.VBComponents.Import sourcePath & file
        End If
        
        file = Dir
    Wend
End Sub
Public Sub SaveSourceFiles() '(destPath As String)
    Set project = Application.VBE.ActiveVBProject
    
    
    
    
End Sub

Sub AddNewCB()
   Set project = Application.VBE.ActiveVBProject
    
   Dim myCommandBar As CommandBar, myCommandBarCtl As CommandBarControl
   On Error GoTo AddNewCB_Err
   Set myCommandBar = project.VBE.CommandBars.Add(Name:="Sample Toolbar", Position:=msoBarFloating)
   myCommandBar.Visible = True
   Set myCommandBarCtl = myCommandBar.Controls.Add(Type:=msoControlButton)
   With myCommandBarCtl
      .Caption = "Button"
      .Style = msoButtonCaption
      .TooltipText = "Display Message Box"
      .OnAction = "=MsgBox ""You pressed a toolbar button!"""
   End With
   Set myCommandBarCtl = myCommandBar.Controls.Add(Type:=msoControlButton)
   With myCommandBarCtl
      .FaceId = 1000
      .Caption = "Toggle Button"
      .TooltipText = "Toggle First Button"
      .OnAction = "=ToggleButton()"
   End With
   Set myCommandBarCtl = myCommandBar.Controls.Add(msoControlComboBox)
   With myCommandBarCtl
      .Caption = "Drop Down"
      .Width = 100
      .AddItem "Create Button", 1
      .AddItem "Remove Button", 2
      .DropDownWidth = 100
      .OnAction = "=AddRemoveButton()"
   End With
   Exit Sub
AddNewCB_Err:
   Debug.Print Err.Number & vbCr & Err.Description
   Exit Sub
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
    
    
    Set project = Application.VBE.ActiveVBProject
    
    ignorePath = destPath & "ignore\"
    'Clean out the ignore folder
    fileSystemHandler.DeleteFile (ignorePath & "*")
    
    
    'Loop through all of the components and export to the ignore folder
    For Each comp In project.VBComponents
        If Left(comp.Name, 5) <> "Sheet" And comp.Name <> "ThisWorkbook" And comp.Name <> "DevTools" And comp.Name <> "ThisDrawing" Then
            'cleanFilePath = destPath & comp.Name & ToFileExtension(comp.Type)
            ignoreFilePath = ignorePath & comp.Name & ToFileExtension(comp.Type)
            comp.Export ignoreFilePath
        End If
    Next
    
    
    'Loop through root directory and find files with no match in ignore folder
    For Each rootFile In fileSystemHandler.GetFolder(destPath).Files
        
        'Check that this file is a type handled by VBA
        If InStr(VBAFileTypes, GetFileExtension(rootFile.Name, True)) > 0 Then
        
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
    
        If InStr(GoodImport, GetFileExtension(ignoreFile.Name, True)) > 0 Then
            'Check that file exists in root, otherwise it should be copied by default
            If Dir(destPath & ignoreFile.Name) <> "" Then
                'Move the file if the CRC32 doesn't match.. otherwise delete it
                MsgBox Hex$(CRC32.CalcCRC32(ignorePath & ignoreFile.Name)) & "i"
                MsgBox Hex$(CRC32.CalcCRC32(destPath & ignoreFile.Name)) & "d"
                If Hex$(CRC32.CalcCRC32(ignorePath & ignoreFile.Name)) <> Hex$(CRC32.CalcCRC32(destPath & ignoreFile.Name)) Then
                    ignoreFile.Copy destPath & ignoreFile.Name, True
                    
                    'Also copy the match FRX file if copying a form
                    If GetFileExtension(ignoreFile.Name, True) = "FRM" Then
                        frxName = Left(ignoreFile.Name, Len(ignoreFile.Name) - 3) & "frx"
                        ignoreFilePath = ignorePath & frxName
                        cleanFilePath = destPath & frxName
                        
                        fileSystemHandler.CopyFile ignoreFilePath, cleanFilePath, True
                        fileSystemHandler.DeleteFile (ignoreFilePath)
                        
                    End If
                    
                    'ignoreFile.Delete
                    
                Else
                    'ignoreFile.Delete
                End If
                
                
            Else
                'Move from ignore to root since file doesn't exist
                ignoreFile.Move (destPath)
                
            
                'Also copy the match FRX file if copying a form
                If GetFileExtension(ignoreFile.Name, True) = "FRM" Then
                    frxName = Left(ignoreFile.Name, Len(ignoreFile.Name) - 3) & "frx"
                    ignoreFilePath = ignorePath & frxName
                    cleanFilePath = destPath & frxName
                    
                    fileSystemHandler.CopyFile ignoreFilePath, cleanFilePath, True
                    fileSystemHandler.DeleteFile (ignoreFilePath)
                    
                End If
                
            End If
            
            
            
        End If
        
        
    Next
    
    
    
    
'
'
'
'
'
'            If ToFileExtension(comp.Type) = ".frm" Then 'Doing this because of craziness regarding .frx file
'
'                'Check if form already exists in source control
'                If Dir(destPath & comp.Name & ToFileExtension(comp.Type)) <> "" Then
'
'                    'If hash of newly exported form (.ign) matches existing form then there is no point in re-exporting
'                    If Not (Hex$(CRC32.CalcCRC32(ignoreFilePath)) = Hex$(CRC32.CalcCRC32(cleanFilePath))) Then
'                        'Re-export files out ... don't bother deleting .ign since they are ignored
'                        'MsgBox "File CRC32 do not match"
'                        comp.Export cleanFilePath
'                    Else
'                        'Do not re-export files since the CRC32 matches
'                        'MsgBox "File CRC32 match"
'
'                    End If
'
'                Else
'                    'Export without checking CRC32 since form file does not exist
'                    comp.Export cleanFilePath
'                End If
'
'            Else
'                comp.Export cleanFilePath
'            End If
'
'
'
'            'MsgBox Hex$(CRC32.CalcCRC32(destPath & comp.Name & ToFileExtension(comp.Type)))
'
'            'MsgBox destPath & comp.Name & ToFileExtension(comp.Type)
'        End If
'    Next
    
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

