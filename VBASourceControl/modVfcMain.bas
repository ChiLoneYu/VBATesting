Attribute VB_Name = "modVfcMain"
Option Explicit

Public Const ccAPPNAME As String = "VBA²VB Form Converter"
Public Const ccAPPVER As String = "0.12"

Private mobjMSForm As VBIDE.VBComponent
Private mobjControl As MSForms.Control

Private mstrSaveAs As String
Private mblnIncludeCode As Boolean
Private mblnShowUnknown As Boolean
Private mintIndent As Integer
Private mintContainer As Integer
Private mintFile As Integer
Private mintUnknownControls As Integer

Public mstrAcadVersion As Integer
Public mstrAcadPlatform As Integer
Public mstrAcadVertical As String

Public Function SelectFormToConvert()
  'START HERE
  frmVfcMain.show vbModal
End Function

Public Sub ProcessForm(strFormName As String, blnProceed As Boolean)
  Dim objControl As MSForms.Control
  Dim strMessage As String
  
  'Setup variables
  Set mobjMSForm = Application.VBE.VBProjects(frmVfcMain.SelectedProjectIndex).VBComponents(strFormName)
  mstrSaveAs = frmVfcMain.SelectedProjectGitIgnoreDirectory & mobjMSForm.Name & ".frmcmp"
  mblnIncludeCode = frmVfcMain.IncludeCode
  mblnShowUnknown = frmVfcMain.ShowUnknown
  mintIndent = 0
  mintContainer = 0
  mintFile = FreeFile
  mintUnknownControls = 0
  
  'Debug report
  Debug.Print "@--- VFC Debug Report ---@"
  Debug.Print "  Form Being Exported: "; mobjMSForm.Name
  Debug.Print "  Filename of New Form: "; mstrSaveAs
  Debug.Print "  Including Code?: "; mblnIncludeCode
  Debug.Print "  Showing Unknown Controls?: "; mblnShowUnknown
  
  'Remove the following IF block if you
  'dont want it to check for overwriting
'  If Dir(mstrSaveAs) <> "" Then
'    If MsgBox(mstrSaveAs & " already exists." & vbCrLf & "Do you want to replace it?", _
'              vbYesNo Or vbExclamation, ccAPPNAME) = vbNo Then
'      blnProceed = False
'    End If
'  End If
  
  If blnProceed Then
    'Convert form
    Open mstrSaveAs For Output As #mintFile
      Call WriteFormHeader
      Call WriteFormProperties
      For Each objControl In mobjMSForm.Designer.Controls
        If objControl.Parent.Name = mobjMSForm.Name Then
          Set mobjControl = objControl
          Call ProcessControl(0)
        End If
      Next objControl
      Call WriteFormFooter
      Call WriteFormCode
    Close #mintFile
    
    'Completion message
    strMessage = "Form Name = " & strFormName & vbCrLf & "Form Conversion completed "
    If mintUnknownControls > 0 Then
      strMessage = strMessage & "with errors." & vbCrLf
      strMessage = strMessage & CStr(mintUnknownControls) & " unknown controls were found." & vbCrLf
      If mblnShowUnknown Then
        strMessage = strMessage & "Please examine the form before use." & vbCrLf
      Else
        strMessage = strMessage & "They have been omitted from the form." & vbCrLf
      End If
    Else
      strMessage = strMessage & "successfully." & vbCrLf
    End If
    strMessage = strMessage & "The new form file was saved to:" & vbCrLf & vbCrLf & mstrSaveAs
    Debug.Print strMessage ', vbInformation, ccAPPNAME
  
  End If
  
  Set mobjControl = Nothing
  Set objControl = Nothing
  Set mobjMSForm = Nothing
End Sub

Private Sub ProcessControl(intContainer As Integer)
  'Debug.Print mobjControl.Name, mobjControl.Parent.Name
  mintContainer = intContainer
  
  If TypeOf mobjControl Is MSForms.Label Then
    Call WriteLabelProperties
  ElseIf TypeOf mobjControl Is MSForms.TextBox Then
    Call WriteTextBoxProperties
  ElseIf TypeOf mobjControl Is MSForms.CheckBox Then
    Call WriteCheckBoxProperties
  ElseIf TypeOf mobjControl Is MSForms.OptionButton Then
    Call WriteOptionButtonProperties
  ElseIf TypeOf mobjControl Is MSForms.CommandButton Then
    Call WriteCommandButtonProperties
  ElseIf TypeOf mobjControl Is MSForms.ToggleButton Then
    Call WriteToggleButtonProperties
  ElseIf TypeOf mobjControl Is MSForms.Image Then
    Call WriteImageProperties
  ElseIf TypeOf mobjControl Is MSForms.ListBox Then
    Call WriteListBoxProperties
  ElseIf TypeOf mobjControl Is MSForms.ComboBox Then
    Call WriteComboBoxProperties
  ElseIf TypeOf mobjControl Is MSForms.ScrollBar Then
    Call WriteScrollBarProperties
  ElseIf TypeOf mobjControl Is MSForms.Frame Then
    Call WriteFrameProperties
  Else
    Call WriteUnknownProperties
  End If
  
End Sub

Private Sub WriteFrameProperties()
  Dim objControl As MSForms.Control
  Dim strCurrentParent As String
  
  Print #mintFile, Indent; "Begin VB.Frame "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  'Write control properties
  Call WriteBackColor(&H8000000F)
  Call WriteCaption
  Call WriteEnabled
  If ParentFontDifferent Then Call WriteFont(mobjControl.Font)
  Call WriteForeColor(&H80000012)
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteLeft
  Call WriteMousePointer
  Call WriteTabIndex
  Call WriteTag
  Call WriteToolTipText
  Call WriteTop
  Call WriteVisible
  Call WriteWidth

  strCurrentParent = mobjControl.Name
  For Each objControl In mobjMSForm.Designer.Controls
    If objControl.Parent.Name = strCurrentParent Then
      Set mobjControl = objControl
      Call ProcessControl(1)
    End If
  Next objControl

  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
  Set objControl = Nothing
End Sub

Private Sub WriteLabelProperties()
  Print #mintFile, Indent; "Begin VB.Label "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  Call WriteAlignmentLabel
  Call WriteAutoSize
  Call WriteBackColor(&H8000000F)
  Call WriteBackStyle
  Call WriteCaption
  Call WriteEnabled
  If ParentFontDifferent Then Call WriteFont(mobjControl.Font)
  Call WriteForeColor(&H80000012)
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteLeft
  Call WriteMousePointer
  Call WriteTabIndex
  Call WriteTag
  Call WriteToolTipText
  Call WriteTop
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteTextBoxProperties()
  Print #mintFile, Indent; "Begin VB.TextBox "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  Call WriteAlignmentLabel
  Call WriteBackColor(&H80000005)
  Call WriteEnabled
  If ParentFontDifferent Then Call WriteFont(mobjControl.Font)
  Call WriteForeColor(&H80000008)
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteHideSelection
  Call WriteLeft
  Call WriteLocked
  Call WriteMaxLength
  Call WriteMousePointer
  Call WriteMultiLine
  Call WritePasswordChar
  Call WriteScrollBars
  Call WriteTabIndex
  Call WriteTabStop
  Call WriteTag
  Call WriteText
  Call WriteToolTipText
  Call WriteTop
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteCheckBoxProperties()
  Print #mintFile, Indent; "Begin VB.CheckBox "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  Call WriteAlignmentCheckBox
  Call WriteBackColor(&H8000000F)
  Call WriteCaption
  Call WriteEnabled
  If ParentFontDifferent Then Call WriteFont(mobjControl.Font)
  Call WriteForeColor(&H80000012)
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteLeft
  Call WriteMousePointer
  Call WriteTabIndex
  Call WriteTabStop
  Call WriteTag
  Call WriteToolTipText
  Call WriteTop
  Call WriteValueCheckBox
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteOptionButtonProperties()
  Print #mintFile, Indent; "Begin VB.OptionButton "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  Call WriteAlignmentCheckBox
  Call WriteBackColor(&H8000000F)
  Call WriteCaption
  Call WriteEnabled
  If ParentFontDifferent Then Call WriteFont(mobjControl.Font)
  Call WriteForeColor(&H80000012)
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteLeft
  Call WriteMousePointer
  Call WriteTabIndex
  Call WriteTabStop
  Call WriteTag
  Call WriteToolTipText
  Call WriteTop
  Call WriteValueOptionButton
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteCommandButtonProperties()
  Print #mintFile, Indent; "Begin VB.CommandButton "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  Call WriteBackColor(&H8000000F)
  Call WriteCancel
  Call WriteCaption
  Call WriteEnabled
  If ParentFontDifferent Then Call WriteFont(mobjControl.Font)
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteLeft
  Call WriteMousePointer
  Call WriteStyleCommandButton
  Call WriteTabIndex
  Call WriteTabStop
  Call WriteTag
  Call WriteToolTipText
  Call WriteTop
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteToggleButtonProperties()
  Print #mintFile, Indent; "Begin VB.CheckBox "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  Call WriteBackColor(&H8000000F)
  Call WriteCaption
  Call WriteEnabled
  If ParentFontDifferent Then Call WriteFont(mobjControl.Font)
  Call WriteForeColor(&H80000012)
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteLeft
  Call WriteMousePointer
  Call WriteStyleToggleButton
  Call WriteTabIndex
  Call WriteTabStop
  Call WriteTag
  Call WriteToolTipText
  Call WriteTop
  Call WriteValueCheckBox
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteImageProperties()
  Print #mintFile, Indent; "Begin VB.Image "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  Call WriteEnabled
  Call WriteHeight
  Call WriteLeft
  Call WriteMousePointer
  Call WriteStretch
  Call WriteTag
  Call WriteToolTipText
  Call WriteTop
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteListBoxProperties()
  Print #mintFile, Indent; "Begin VB.ListBox "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  Call WriteBackColor(&H80000005)
  Call WriteColumns
  Call WriteEnabled
  If ParentFontDifferent Then Call WriteFont(mobjControl.Font)
  Call WriteForeColor(&H80000008)
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteIntegralHeight
  Call WriteLeft
  Call WriteMousePointer
  Call WriteMultiSelect
  Call WriteStyleListBox
  Call WriteTabIndex
  Call WriteTabStop
  Call WriteTag
  Call WriteToolTipText
  Call WriteTop
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteComboBoxProperties()
  Print #mintFile, Indent; "Begin VB.ComboBox "; mobjControl.Name
  mintIndent = mintIndent + 1
  
  Call WriteBackColor(&H80000005)
  Call WriteEnabled
  If ParentFontDifferent Then Call WriteFont(mobjControl.Font)
  Call WriteForeColor(&H80000008)
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteLeft
  Call WriteLocked
  Call WriteMousePointer
  Call WriteStyleComboBox
  Call WriteTabIndex
  Call WriteTabStop
  Call WriteTag
  Call WriteText
  Call WriteToolTipText
  Call WriteTop
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteScrollBarProperties()
  If mobjControl.Orientation = fmOrientationAuto Then
    If mobjControl.Width > mobjControl.Height Then
      Print #mintFile, Indent; "Begin VB.HScrollBar "; mobjControl.Name
    Else
      Print #mintFile, Indent; "Begin VB.VScrollBar "; mobjControl.Name
    End If
  ElseIf mobjControl.Orientation = fmOrientationVertical Then
    Print #mintFile, Indent; "Begin VB.VScrollBar "; mobjControl.Name
  Else 'If mobjControl.Orientation = fmOrientationHorizontal Then
    Print #mintFile, Indent; "Begin VB.HScrollBar "; mobjControl.Name
  End If
  mintIndent = mintIndent + 1
  
  Call WriteEnabled
  Call WriteHeight
  Call WriteHelpContextID
  Call WriteLargeChange
  Call WriteLeft
  Call WriteMaxScrollBar
  Call WriteMinScrollBar
  Call WriteMousePointer
  Call WriteSmallChange
  Call WriteTabIndex
  Call WriteTabStop
  Call WriteTag
  Call WriteTop
  Call WriteValueScrollBar
  Call WriteVisible
  Call WriteWidth
  
  mintIndent = mintIndent - 1
  Print #mintFile, Indent; "End"
End Sub

Private Sub WriteUnknownProperties()
  
  mintUnknownControls = mintUnknownControls + 1
  
  If mblnShowUnknown Then
    'Show the unknown control on the converted form as a red label
    Print #mintFile, Indent & "Begin VB.Label " & mobjControl.Name
    mintIndent = mintIndent + 1
    
    Print #mintFile, FormatProperty("Alignment"); "2  'Center"
    Print #mintFile, FormatProperty("Appearance"); "0  'Flat"
    Print #mintFile, FormatProperty("BackColor"); "&H000000FF&"
    Print #mintFile, FormatProperty("BorderStyle"); "1  'Fixed Single"
    Print #mintFile, FormatProperty("Caption"); FormatString(mobjControl.Name & " - " & TypeName(mobjControl))
    Call WriteHeight
    Call WriteLeft
    Call WriteTop
    Call WriteWidth
  
    mintIndent = mintIndent - 1
    Print #mintFile, Indent & "End"
  End If
  
  'Print the unknown control to the immediate window
  Debug.Print "@--- Unknown Control Found ---@"
  Debug.Print "  Control Name: "; mobjControl.Name
  Debug.Print "  Control Type: "; TypeName(mobjControl)
  
End Sub

Private Sub WriteAlignmentLabel()
  'ALIGNMENT
  If mobjControl.TextAlign <> fmTextAlignLeft Then
    If mobjControl.TextAlign = fmTextAlignRight Then
      Print #mintFile, FormatProperty("Alignment"); "1  'Right Jusify"
    ElseIf mobjControl.TextAlign = fmTextAlignCenter Then
      Print #mintFile, FormatProperty("Alignment"); "2  'Center"
    End If
  End If
End Sub

Private Sub WriteAlignmentCheckBox()
  'ALIGNMENT
  If mobjControl.Alignment <> fmAlignmentRight Then
    Print #mintFile, FormatProperty("Alignment"); "1  'Right Jusify"
  End If
End Sub

Private Sub WriteAutoSize()
  'AUTOSIZE
  If mobjControl.AutoSize = True Then
    Print #mintFile, FormatProperty("AutoSize"); "-1  'True"
  End If
End Sub

Private Sub WriteBackColor(lngDefault As Long)
  'BACKCOLOR
  If mobjControl.BackColor <> lngDefault Then
    Print #mintFile, FormatProperty("BackColor"); FormatHex(mobjControl.BackColor)
  End If
End Sub

Private Sub WriteBackStyle()
  'BACKSTYLE
  If mobjControl.BackStyle <> fmBackStyleOpaque Then
    Print #mintFile, FormatProperty("BackStyle"); "0  'Transparent"
  End If
End Sub

Private Sub WriteCancel()
  'CANCEL
  If mobjControl.Cancel = True Then
    Print #mintFile, FormatProperty("Cancel"); "-1  'True"
  End If
End Sub

Private Sub WriteCaption()
  Dim strCaption As String
  Dim strChar As String
  Dim strTemp As String
  Dim intPos As Integer
  
  strCaption = mobjControl.Caption
  If strCaption <> "" Then
    'Find &'s and replace with &&'s
    For intPos = 1 To Len(strCaption)
      strChar = Mid(strCaption, intPos, 1)
      If strChar = "&" Then
        strTemp = strTemp & "&"
      End If
      strTemp = strTemp & strChar
    Next intPos
    'Add mnemonic
    If TypeOf mobjControl Is MSForms.Label Or _
       TypeOf mobjControl Is MSForms.CheckBox Or _
       TypeOf mobjControl Is MSForms.OptionButton Or _
       TypeOf mobjControl Is MSForms.CommandButton Or _
       TypeOf mobjControl Is MSForms.ToggleButton Then
      strChar = mobjControl.Accelerator
      If strChar <> "" Then
        intPos = InStr(1, strTemp, strChar, vbBinaryCompare)
        If intPos > 0 Then
          strTemp = Left(strTemp, intPos - 1) & "&" & Mid(strTemp, intPos)
        End If
      End If
    End If
    'CAPTION
    Print #mintFile, FormatProperty("Caption"); FormatString(strTemp)
  End If
End Sub

Private Sub WriteColumns()
  'COLUMNS
  If mobjControl.ColumnCount > 1 Then
    Print #mintFile, FormatProperty("Columns"); CStr(mobjControl.ColumnCount - 1)
  End If
End Sub

Private Sub WriteDefault()
  'DEFAULT
  If mobjControl.Default = True Then
    Print #mintFile, FormatProperty("Default"); "-1  'True"
  End If
End Sub

Private Sub WriteEnabled()
  'ENABLED
  If mobjControl.Enabled = False Then
    Print #mintFile, FormatProperty("Enabled"); "0  'False"
  End If
End Sub

Private Sub WriteForeColor(lngDefault As Long)
  'FORECOLOR
  If mobjControl.ForeColor <> lngDefault Then
    Print #mintFile, FormatProperty("ForeColor"); FormatHex(mobjControl.ForeColor)
  End If
End Sub

Private Sub WriteHeight()
  'HEIGHT
  Print #mintFile, FormatProperty("Height"); CStr(mobjControl.Height * 20)
End Sub

Private Sub WriteHelpContextID()
  'HELPCONTEXTID
  If mobjControl.HelpContextID <> 0 Then
    Print #mintFile, FormatProperty("HelpContextID"); CStr(mobjControl.HelpContextID)
  End If
End Sub

Private Sub WriteHideSelection()
  'HIDESELECTION
  If mobjControl.HideSelection = False Then
    Print #mintFile, FormatProperty("HideSelection"); "0  'False"
  End If
End Sub

Private Sub WriteIntegralHeight()
  'INTEGRALHEIGHT
  If mobjControl.IntegralHeight = False Then
    Print #mintFile, FormatProperty("IntegralHeight"); "0  'False"
  End If
End Sub

Private Sub WriteLargeChange()
  'LARGECHANGE
  If mobjControl.LargeChange <> 1 Then
    Print #mintFile, FormatProperty("LargeChange"); CStr(mobjControl.LargeChange)
  End If
End Sub

Private Sub WriteLeft()
  Dim lngOffset As Long
  If mintContainer = 1 Then lngOffset = 30
  'LEFT
  Print #mintFile, FormatProperty("Left"); CStr(mobjControl.Left * 20 + lngOffset)
End Sub

Private Sub WriteLocked()
  'LOCKED
  If mobjControl.Locked = True Then
    Print #mintFile, FormatProperty("Locked"); "-1  'True"
  End If
End Sub

Private Sub WriteMaxScrollBar()
  'MAX
  If mobjControl.Max <> 32767 Then
    Print #mintFile, FormatProperty("Max"); CStr(mobjControl.Max)
  End If
End Sub

Private Sub WriteMaxLength()
  'MAXLENGTH
  If mobjControl.MaxLength <> 0 Then
    Print #mintFile, FormatProperty("MaxLength"); CStr(mobjControl.MaxLength)
  End If
End Sub

Private Sub WriteMinScrollBar()
  'MIN
  If mobjControl.Min <> 0 Then
    Print #mintFile, FormatProperty("Min"); CStr(mobjControl.Min)
  End If
End Sub

Private Sub WriteMousePointer()
  'MOUSEPOINTER
  If mobjControl.MousePointer <> fmMousePointerDefault Then
    Print #mintFile, FormatProperty("MousePointer"); CStr(mobjControl.MousePointer)
  End If
End Sub

Private Sub WriteMultiLine()
  'MULTILINE
  If mobjControl.MultiLine = True Then
    Print #mintFile, FormatProperty("MultiLine"); "-1  'True"
  End If
End Sub

Private Sub WriteMultiSelect()
  'MULTISELECT
  If mobjControl.MultiSelect <> fmMultiSelectSingle Then
    If mobjControl.MultiSelect = fmMultiSelectMulti Then
      Print #mintFile, FormatProperty("MultiSelect"); "1  'Simple"
    ElseIf mobjControl.MultiSelect = fmMultiSelectExtended Then
      Print #mintFile, FormatProperty("MultiSelect"); "2  'Extended"
    End If
  End If
End Sub

Private Sub WritePasswordChar()
  'PASSWORDCHAR
  If mobjControl.PasswordChar <> "" Then
    Print #mintFile, FormatProperty("PasswordChar"); FormatString(mobjControl.PasswordChar)
  End If
End Sub

Private Sub WriteScrollBars()
  'SCROLLBARS
  If mobjControl.ScrollBars <> fmScrollBarsNone Then
    If mobjControl.ScrollBars = fmScrollBarsHorizontal Then
      Print #mintFile, FormatProperty("ScrollBars"); "1  'Horizontal"
    ElseIf mobjControl.ScrollBars = fmScrollBarsVertical Then
      Print #mintFile, FormatProperty("ScrollBars"); "2  'Vertical"
    ElseIf mobjControl.ScrollBars = fmScrollBarsBoth Then
      Print #mintFile, FormatProperty("ScrollBars"); "3  'Both"
    End If
  End If
End Sub

Private Sub WriteSmallChange()
  'SMALLCHANGE
  If mobjControl.SmallChange <> 1 Then
    Print #mintFile, FormatProperty("SmallChange"); CStr(mobjControl.SmallChange)
  End If
End Sub

Private Sub WriteStretch()
  'STRETCH
  If mobjControl.PictureSizeMode <> fmPictureSizeModeClip Then
    Print #mintFile, FormatProperty("Stretch"); "-1  'True"
  End If
End Sub

Private Sub WriteStyleComboBox()
  'STYLE
  If mobjControl.Style <> fmStyleDropDownCombo Then
    Print #mintFile, FormatProperty("Style"); "2  'Dropdown List"
  End If
End Sub

Private Sub WriteStyleCommandButton()
  'STYLE
  If mobjControl.Picture <> 0 Then
    Print #mintFile, FormatProperty("Style"); "1  'Graphical"
  End If
End Sub

Private Sub WriteStyleListBox()
  'STYLE
  If mobjControl.ListStyle <> fmListStylePlain Then
    Print #mintFile, FormatProperty("Style"); "1  'Checkbox"
  End If
End Sub

Private Sub WriteStyleToggleButton()
  'STYLE
  Print #mintFile, FormatProperty("Style"); "1  'Graphical"
End Sub

Private Sub WriteTabIndex()
  'TABINDEX
  Print #mintFile, FormatProperty("TabIndex"); CStr(mobjControl.TabIndex)
End Sub

Private Sub WriteTabStop()
  'TABSTOP
  If mobjControl.TabStop = False Then
    Print #mintFile, FormatProperty("TabStop"); "0  'False"
  End If
End Sub

Private Sub WriteTag()
  'TAG
  If mobjControl.Tag <> "" Then
    Print #mintFile, FormatProperty("Tag"); FormatString(mobjControl.Tag)
  End If
End Sub

Private Sub WriteText()
  'TEXT
  If mobjControl.Text <> "" Then
    Print #mintFile, FormatProperty("Text"); Left(FormatString(mobjControl.Text), 2047)
  End If
End Sub

Private Sub WriteToolTipText()
  'TOOLTIPTEXT
  If mobjControl.ControlTipText <> "" Then
    Print #mintFile, FormatProperty("ToolTipText"); FormatString(mobjControl.ControlTipText)
  End If
End Sub

Private Sub WriteTop()
  Dim lngOffset As Long
  If mintContainer = 1 Then lngOffset = 120
  'TOP
  Print #mintFile, FormatProperty("Top"); CStr(mobjControl.Top * 20 + lngOffset)
End Sub

Private Sub WriteValueCheckBox()
  'VALUE
  If mobjControl.Value = True Then
    Print #mintFile, FormatProperty("Value"); "1  'Checked"
  ElseIf IsNull(mobjControl.Value) Then
    Print #mintFile, FormatProperty("Value"); "2  'Grayed"
  End If
End Sub

Private Sub WriteValueOptionButton()
  'VALUE
  If mobjControl.Value = True Then
    Print #mintFile, FormatProperty("Value"); "-1  'True"
  End If
End Sub

Private Sub WriteValueScrollBar()
  'VALUE
  If mobjControl.Value <> 0 Then
    Print #mintFile, FormatProperty("Value"); CStr(mobjControl.Value)
  End If
End Sub

Private Sub WriteVisible()
  'VISIBLE
  If mobjControl.Visible = False Then
    Print #mintFile, FormatProperty("Visible"); "0  'False"
  End If
End Sub

Private Sub WriteWidth()
  'WIDTH
  Print #mintFile, FormatProperty("Width"); CStr(mobjControl.Width * 20)
End Sub

Private Sub WriteFormHeader()
  Print #mintFile, "VERSION 5.00"
  Print #mintFile, "Begin VB.Form "; mobjMSForm.Name
End Sub

Private Sub WriteFormProperties()
  Dim objUserForm As MSForms.UserForm
  
  Set objUserForm = mobjMSForm.Designer
  mintIndent = 1
  
  'BACKCOLOR
  If objUserForm.BackColor <> &H8000000F Then
    Print #mintFile, FormatProperty("BackColor"); FormatHex(objUserForm.BackColor)
  End If
  
  'BORDERSTYLE - Set to fixed single
  'This property although not included in VBA causes the VB form to act
  'like a VBA form, make changes to this property after importing
  Print #mintFile, FormatProperty("BorderStyle"); "1  'Fixed Single"
  
  'CAPTION
  If objUserForm.Caption <> "" Then
    Print #mintFile, FormatProperty("Caption"); FormatString(objUserForm.Caption)
  End If
  
  'CLIENTHEIGHT
  Print #mintFile, FormatProperty("ClientHeight"); CStr(objUserForm.InsideHeight * 20)
  
  '-CLIENTLEFT
  Print #mintFile, FormatProperty("ClientLeft"); CStr(mobjMSForm.Properties("Left").Value * 20 + 45)
  
  '-CLIENTTOP
  Print #mintFile, FormatProperty("ClientTop"); CStr(mobjMSForm.Properties("Top").Value * 20 + 330)
  
  'CLIENTWIDTH
  Print #mintFile, FormatProperty("ClientWidth"); CStr(objUserForm.InsideWidth * 20)
  
  'ENABLED
  If objUserForm.Enabled = False Then
    Print #mintFile, FormatProperty("Enabled"); "0"
  End If
  
  'FONT
  Call WriteFont(objUserForm.Font)
  
  'FORECOLOR
  If objUserForm.ForeColor <> &H80000012 Then
    Print #mintFile, FormatProperty("ForeColor"); FormatHex(objUserForm.ForeColor)
  End If
  
  '-HELPCONTEXTID
  If mobjMSForm.Properties("HelpContextID").Value <> 0 Then
    Print #mintFile, FormatProperty("HelpContextID"); CStr(mobjMSForm.Properties("HelpContextID").Value)
  End If
  
  'MAXBUTTON
  'This property although not included in VBA causes the VB form to act
  'like a VBA form, make changes to this property after importing
  Print #mintFile, FormatProperty("MaxButton"); "0   'False"
  
  'MINBUTTON
  'This property although not included in VBA causes the VB form to act
  'like a VBA form, make changes to this property after importing
  Print #mintFile, FormatProperty("MinButton"); "0   'False"
  
  'MOUSEPOINTER
  If objUserForm.MousePointer <> fmMousePointerDefault Then
    Print #mintFile, FormatProperty("MousePointer") & CStr(objUserForm.MousePointer)
  End If
  
  'SCALEHEIGHT
  Print #mintFile, FormatProperty("ScaleHeight") & CStr(objUserForm.InsideHeight * 20)
  
  'SCALEWIDTH
  Print #mintFile, FormatProperty("ScaleWidth") & CStr(objUserForm.InsideWidth * 20)
  
  '-STARTUPPOSITION
  Select Case mobjMSForm.Properties("StartUpPosition").Value
    Case 0
      Print #mintFile, FormatProperty("StartUpPosition"); "0  'Manual"
    Case 1
      Print #mintFile, FormatProperty("StartUpPosition"); "1  'CenterOwner"
    Case 2
      Print #mintFile, FormatProperty("StartUpPosition"); "2  'CenterScreen"
  End Select
  
  '-TAG
  If mobjMSForm.Properties("Tag").Value <> "" Then
    Print #mintFile, FormatProperty("Tag"); FormatString(mobjMSForm.Properties("Tag").Value)
  End If
  
  '-WHATSTHISBUTTON
  If mobjMSForm.Properties("WhatsThisButton").Value = True Then
    Print #mintFile, FormatProperty("WhatsThisButton"); "-1  'True"
  End If
  
  '-WHATSTHISHELP
  If mobjMSForm.Properties("WhatsThisHelp").Value = True Then
    Print #mintFile, FormatProperty("WhatsThisHelp"); "-1  'True"
  End If
  
  Set objUserForm = Nothing
End Sub

Private Sub WriteFormFooter()
  Print #mintFile, "End"
  Print #mintFile, "Attribute VB_Name = """ & mobjMSForm.Name & """"
  'TODO: figure out if these are the right settings to use
  Print #mintFile, "Attribute VB_GlobalNameSpace = False"
  Print #mintFile, "Attribute VB_Creatable = False"
  Print #mintFile, "Attribute VB_PredeclaredId = True"
  Print #mintFile, "Attribute VB_Exposed = False"
End Sub

Private Sub WriteFormCode()
  Dim lngLine As Long
  If mblnIncludeCode Then
    For lngLine = 1 To mobjMSForm.CodeModule.CountOfLines
      Print #mintFile, mobjMSForm.CodeModule.Lines(lngLine, 1)
    Next lngLine
  End If
End Sub

Private Sub WriteFont(objFont As stdole.StdFont)
  Dim strProperty As String
  Print #mintFile, Indent & "BeginProperty Font"
  mintIndent = mintIndent + 1
  Print #mintFile, FormatProperty("Name") & FormatString(objFont.Name)
  Print #mintFile, FormatProperty("Size") & Replace(CStr(objFont.Size), ",", ".")
  Print #mintFile, FormatProperty("Charset") & CStr(objFont.Charset)
  Print #mintFile, FormatProperty("Weight") & IIf(objFont.Bold, "700", "400")
  Print #mintFile, FormatProperty("Underline") & IIf(objFont.Underline, "-1  'True", "0  'False")
  Print #mintFile, FormatProperty("Italic") & IIf(objFont.Italic, "-1  'True", "0  'False")
  Print #mintFile, FormatProperty("Strikethrough") & IIf(objFont.Strikethrough, "-1  'True", "0  'False")
  mintIndent = mintIndent - 1
  Print #mintFile, Indent & "EndProperty"
End Sub

Private Function ParentFontDifferent() As Boolean
  If mobjControl.Parent.Font.Name <> mobjControl.Font.Name Then
    ParentFontDifferent = True
  ElseIf mobjControl.Parent.Font.Size <> mobjControl.Font.Size Then
    ParentFontDifferent = True
  ElseIf mobjControl.Parent.Font.Charset <> mobjControl.Font.Charset Then
    ParentFontDifferent = True
  ElseIf mobjControl.Parent.Font.Weight <> mobjControl.Font.Weight Then
    ParentFontDifferent = True
  ElseIf mobjControl.Parent.Font.Underline <> mobjControl.Font.Underline Then
    ParentFontDifferent = True
  ElseIf mobjControl.Parent.Font.Italic <> mobjControl.Font.Italic Then
    ParentFontDifferent = True
  ElseIf mobjControl.Parent.Font.Strikethrough <> mobjControl.Font.Strikethrough Then
    ParentFontDifferent = True
  End If
End Function

Private Function FormatProperty(strPropName As String) As String
  If Len(strPropName) < 16 Then
    FormatProperty = Indent & strPropName & Space(16 - Len(strPropName)) & "=   "
  Else
    FormatProperty = Indent & strPropName & " =   "
  End If
End Function

Private Function FormatString(strValue As String) As String
  Dim strChar As String
  Dim strTemp As String
  Dim intPos As Integer
  
  For intPos = 1 To Len(strValue)
    strChar = Mid(strValue, intPos, 1)
    If Asc(strChar) = 34 Then
      strTemp = strTemp & Chr(34)
    End If
    If strChar = vbCr Or strChar = vbLf Then
      strTemp = strTemp & "_"
    Else
      strTemp = strTemp & strChar
    End If
  Next intPos
  
  FormatString = Chr(34) & strTemp & Chr(34)
End Function

Private Function FormatHex(lngValue As Long) As String
  FormatHex = "&H" & Hex(lngValue) & "&"
End Function

Private Function Indent() As String
  Indent = Space(mintIndent * 3)
End Function




'////////////////////////////////////////////////////

Public Function FindAcadVersionAndPlatformAndVertical()
    'get the version by checking the title bar
    Dim titleBarSplited() As String
    titleBarSplited = Split(Application.Caption, " ")
    Dim i As Integer
    mstrAcadVertical = ""
    For i = 0 To UBound(titleBarSplited)
        'first will find the vertical name,
        'then find the version number (year) and then exit for
        If (IsNumeric(titleBarSplited(i))) Then
            mstrAcadVersion = CInt(titleBarSplited(i))
            Exit For
        ElseIf (titleBarSplited(i) <> "AutoCAD") Then
            'vertical names can have more than 1 word
            mstrAcadVertical = mstrAcadVertical + titleBarSplited(i)
        End If
    Next

    'get the if is 32 or 64 bit by checking the PLATFORM variable
    Dim platform As String
    platform = Application.ActiveDocument.GetVariable("PLATFORM")
    If (InStr(1, platform, "x86") > 0) Then
        mstrAcadPlatform = 32
    Else
        mstrAcadPlatform = 64
    End If
    Debug.Print "AutoCAD " & mstrAcadVersion & " " & mstrAcadVertical & " " & mstrAcadPlatform & " bit"
End Function

'Export all files in a VBA project and create a VB6 project to wrap them
Public Function ProcessProject()
    Dim proj As VBIDE.VBProject
    Dim comp As VBIDE.VBComponent
    Dim strFilePath As String
    
    'Grab selected project and store its filepath to use while processing
    Set proj = Application.VBE.VBProjects(frmVfcMain.SelectedProjectIndex)
    strFilePath = frmVfcMain.SelectedProjectGitIgnoreDirectory
     
    'We'll be creating a minimal VB6 project file to wrap together the exported modules.
    'strTxt contains the text for the project file
    Dim strTxt As String
    
    strTxt = "Type=OleDll" & vbCrLf
'
'    'If user wants us to add ObjectDBX Type Library
'    If frmVfcMain.chkAddDbx = True Then
'        'Explicitly add ObjectDBX Type Library as first TLB reference in project file. There are two reasons we do this:
'        '1. Most VBA projects won't reference this (its not needed in VBA), but most VB6/.NET projects will.
'        '     (And it doesn't matter if there are duplicates).
'        '2. Adding it first ensures the Visual Studio conversion Wizard recognizes (for example) AcadLine as coming from the
'        '     Autodesk.AutoCAD.Interop.Common namespace, and not Autodesk.AutoCAD.Interop namespace. (Less editing for you later).
'        Call FindAcadVersionAndPlatformAndVertical
'        If (mstrAcadPlatform = 32) Then
'            Select Case mstrAcadVersion
'                Case 2010 To 2012
'                    strTxt = strTxt & "Reference=*\G{9F83C3E8-AAA3-4B0B-A256-F0DF98DA74BC}#1.0#0#C:\Program Files\Common Files\Autodesk Shared\axdb18enu.tlb#AXDBLib" & vbCrLf
'                Case 2007 To 2009
'                    strTxt = strTxt & "Reference=*\G{11A32D00-9E89-4C16-82CB-629DEBA56AE2}#1.0#0#C:\Program Files\Common Files\Autodesk Shared\axdb17enu.tlb#AXDBLib" & vbCrLf
'                Case Else
'                    MsgBox "Unfortunately this AutoCAD version is not supported", vbCritical
'            End Select
'        Else
'            Select Case mstrAcadVersion
'                Case 2010 To 2012
'                    strTxt = strTxt & "Reference=*\G{FFC2A8DB-A497-4087-941C-C5B5462237EA}#1.0#0#C:\Program Files\Common Files\Autodesk Shared\axdb18enu.tlb#AXDBLib" & vbCrLf
'                Case 2007 To 2009
'                    strTxt = strTxt & "Reference=*\G{B789BF0E-B4A5-46B2-A8FE-D8CE0DA25E63}#1.0#0#C:\Program Files\Common Files\Autodesk Shared\axdb17enu.tlb#AXDBLib" & vbCrLf
'                Case Else
'                    MsgBox "Unfortunately this AutoCAD version is not supported", vbCritical
'            End Select
'        End If
'    End If
    
    'add aditional references specific for verticals
    Select Case mstrAcadVertical
        Case "Civil3D"
            If (MsgBox("This VBA project is running on Civil 3D. Would you like to include aditional references? (AecXUIBase)", vbYesNo) = vbYes) Then
                Select Case mstrAcadVersion
                    Case 2009 'only 32 bit
                        strTxt = strTxt & "Reference=*\G{E7BBE100-BF69-431B-9153-1FF0DEF8F613}#5.7#0#C:\Program Files\Common Files\Autodesk Shared\AecXUIBase57.tlb#AecXUIBase" & vbCrLf
                    Case 2010 'only 32 bit
                        strTxt = strTxt & "Reference=*\G{E7BCE100-BF69-431B-9153-1FF0DEF8F613}#6.0#0#C:\Program Files\Common Files\Autodesk Shared\AecXUIBase60.tlb#AecXUIBase" & vbCrLf
                    Case 2011
                        If (mstrAcadPlatform = 32) Then
                            'TODO
                        Else '64 bits
                            'TODO
                        End If
                    Case Else
                        MsgBox "Unfortunately this Civil 3D version is not supported.", vbCritical
                End Select
            End If
    End Select
       
    Dim ref As VBIDE.Reference
    
    For Each ref In proj.References
        strTxt = strTxt & "Reference=*\G" & ref.GUID & "#" & ref.Major & "." & ref.Minor & "#0#" & ref.FullPath & "#" & ref.Name & vbCrLf
    Next
    
    
    'Iterate each component in project and export it with the right file extension
    'Note special processing for UserForms, which can't be migrated bby .NET Migration Wizard
    '  (have to convert them to VB6 Forms, which is most of the code in the redt of this project)
    For Each comp In proj.VBComponents
        If comp.Type = vbext_ct_MSForm Then
            'Process a form
            Call ProcessForm(comp.Name, True)
            comp.Export (strFilePath & comp.Name & ".frm")
            strTxt = strTxt & "Form=" & comp.Name & ".frmcmp" & vbCrLf
        ElseIf comp.Type = vbext_ct_ClassModule Then
            'Process a Class module
            comp.Export (strFilePath & comp.Name & ".cls")
            strTxt = strTxt & "Class=" & comp.Name & "; " & comp.Name & ".cls" & vbCrLf
        ElseIf comp.Type = vbext_ct_StdModule Then
            'Process a module
            comp.Export (strFilePath & comp.Name & ".bas")
            strTxt = strTxt & "Module=" & comp.Name & "; " & comp.Name & ".bas" & vbCrLf
        ElseIf comp.Type = vbext_ct_Document And comp.Name = "ThisDrawing" Then
            'Process ThisDrawing module (which is exported as a class module)
            comp.Export (strFilePath & comp.Name & ".cls")
            strTxt = strTxt & "Class=" & comp.Name & "; " & comp.Name & ".cls" & vbCrLf
        End If
    Next
    
    'Save project file
    'This is a minimal VB6 project -
    'Remember we're not creating a project to use in VB6, we just want something we can
    '  give to the .NET Migration Wizard.
    
    Open strFilePath & proj.Name & ".vbp" For Output As #42
    Print #42, strTxt
    Close #42
    
    Debug.Print "Project export finished" & vbCrLf & _
        "Project is in folder:" & vbCrLf & _
        strFilePath
End Function


'Public Function PostProcessDotNetProject()
'
'    'If we've selected a project, then we take its directory as our default directory
'    If frmVfcMain.SelectedProjectFilename <> "" Then
'        frmVfcMain.ctrlFileDia.InitDir = frmVfcMain.SelectedProjectFilename
'    End If
'
'    'Ask user to select a file
'    On Error GoTo TheEnd
'    frmVfcMain.ctrlFileDia.DialogTitle = "Select vbproj file to post-process ..."
'    frmVfcMain.ctrlFileDia.DefaultExt = ".vbproj"
'    frmVfcMain.ctrlFileDia.Filter = ".vbproj"
'    frmVfcMain.ctrlFileDia.CancelError = True
'    frmVfcMain.ctrlFileDia.ShowOpen
'    On Error GoTo 0
'
'    'Open document .vbproj file
'    Dim doc As New DOMDocument
'    Dim node As IXMLDOMNode
'
'    If Dir(frmVfcMain.ctrlFileDia.fileName) = "" Then Exit Function
'
'    'Open XML vbproj file.
'    doc.Load frmVfcMain.ctrlFileDia.fileName
'    doc.async = False
'
'    'Set option to launch AutoCAD when we 'F5' debug (we add this to every configuration)
'    '(Note - Normally, VS adds this setting to your vbproj.user file, but it works if you modify the .vbproj file and delete any vbproj.user file that existed before you made the change.
'    '  You can also run this post-processing routine on the .vbproj.user file).
'    For Each node In doc.selectNodes("//PropertyGroup[@Condition]")
'        Dim newNode As IXMLDOMNode
'        Set newNode = node.appendChild(doc.createNode(NODE_ELEMENT, "StartAction", doc.firstChild.namespaceURI))
'        newNode.Text = "Program"
'        Set newNode = node.appendChild(doc.createNode(NODE_ELEMENT, "StartProgram", doc.firstChild.namespaceURI))
'        '*** Change pathname to match the installed location of AutoCAD 20XX on your machine ***
'        Dim verticalPath As String
'        Select Case mstrAcadVertical
'            Case "Civil3D"
'                verticalPath = "Civil 3D "
'        End Select
'        newNode.Text = "C:\Program Files\AutoCAD " & verticalPath & mstrAcadVersion & "\acad.exe"
'
'        'Try remove the x64 bit tags of projects created with Visual Basic Express
'        'http://msdn.microsoft.com/library/we1f72fb.aspx
'        Dim subNode As IXMLDOMNode
'        Dim subNodeConstant As IXMLDOMNode
'        For Each subNode In node.childNodes
'            If (subNode.nodeName = "PlatformTarget") Then
'                subNode.Text = "AnyCPU"
'            ElseIf (subNode.nodeName = "DefineConstants") Then
'                Set subNodeConstant = subNode
'            End If
'        Next
'        Call node.removeChild(subNodeConstant)
'    Next
'
'
'    'We're using VB Express 2008 and AutoCAD 2010, so we want to target Framework 3.5
'    '*** Comment this out if you want to target Framework 2.0 ***
'    Dim nodes As IXMLDOMNodeList
'    Dim newRefNode As IXMLDOMNode
'    Set nodes = doc.selectNodes("//PropertyGroup[ProjectType]")
'    Set node = nodes.Item(0)
'    If Not node Is Nothing Then
'        Set newRefNode = doc.createNode(NODE_ELEMENT, "TargetFrameworkVersion", doc.firstChild.namespaceURI)
'        newRefNode.Text = "v3.5"
'        node.appendChild newRefNode
'    End If
'
'
'    'Add references to acmgd.dll and acdbmgd.dll
'    '*** Edit the text we add below for different DLL versions ***
'    '(Easiest way to find text is to add references manually and then open up the vbproj file in notepad).
'    'We reference the one's from the ObjectARX SDK (because these are better) if we can find them,
'    'otherwise we reference the ones installed with AutoCAD (assuming default install location).
'
'    Dim strRefPath As String
'    strRefPath = frmVfcMain.txtARXSDKLocation
'    If Dir(frmVfcMain.txtARXSDKLocation, vbDirectory) = "" Then
'        'If textbox is empty, assume AutoCAD is installed in default location and use DLLs installed there. (Not ideal).
'        strRefPath = "C:\Program Files\AutoCAD " & verticalPath & mstrAcadVersion & "\"
'    End If
'
'    Set nodes = doc.selectNodes("//ItemGroup[Reference]")
'    Set node = nodes.Item(0)
'    If Not node Is Nothing Then
'
'        Set newRefNode = doc.createNode(NODE_ELEMENT, "Reference", doc.firstChild.namespaceURI)
'        Dim NewAttNode As IXMLDOMNode
'        Set NewAttNode = newRefNode.Attributes.setNamedItem(doc.createAttribute("Include"))
'        NewAttNode.Text = "acmgd, Version=18.0.0.0, Culture=neutral, processorArchitecture=x86"
'        Dim subRefNode As IXMLDOMNode
'        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "SpecificVersion", doc.firstChild.namespaceURI))
'        subRefNode.Text = "False"
'        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "HintPath", doc.firstChild.namespaceURI))
'        subRefNode.Text = strRefPath & "acmgd.dll"
'        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "Private", doc.firstChild.namespaceURI))
'        subRefNode.Text = "False"
'        node.appendChild newRefNode
'
'        Set newRefNode = doc.createNode(NODE_ELEMENT, "Reference", doc.firstChild.namespaceURI)
'        Set NewAttNode = newRefNode.Attributes.setNamedItem(doc.createAttribute("Include"))
'        NewAttNode.Text = "acdbmgd, Version=18.0.0.0, Culture=neutral, processorArchitecture=x86"
'        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "SpecificVersion", doc.firstChild.namespaceURI))
'        subRefNode.Text = "False"
'        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "HintPath", doc.firstChild.namespaceURI))
'        subRefNode.Text = strRefPath & "acdbmgd.dll"
'        Set subRefNode = newRefNode.appendChild(doc.createNode(NODE_ELEMENT, "Private", doc.firstChild.namespaceURI))
'        subRefNode.Text = "False"
'        node.appendChild newRefNode
'
'End If
'
'    'Save new .vbproj file
'    doc.Save frmVfcMain.ctrlFileDia.fileName
'    Set doc = Nothing
'
'    MsgBox "Finished processing .vbproj file", vbInformation, "VBA Converter"
'
'    Exit Function
'
'TheEnd:
'Err.Clear
'End Function




