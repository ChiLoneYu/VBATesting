VBA to VB6 project converter
============================

This macro is an extension of the VBA2VB Form Converter written by Leslie Lowe to migrate an entire VBA project instead of just the VBA UserForms. The bulk of the code is for the Forms conversion, and is taken straight from the original (with Leslie's kind permission). The original version Form Converter is included with this version.

To convert a VBA project:

1. Load the converter macro project and the project(s) you want to convert into AutoCAD (using VBALOAD or VBAMAN commands).
2. Run the macro "ThisDrawing.StartVBConversion" from the converter project (via the VBARUN command, for example).
3. Select the project you want to convert in the list on the form and click the Convert button
4. The exported VB6 project is saved to a 'VB6Conv' subfolder of the folder where the original VBA project was located.

To post-process a vb.NET vbproj file:

1. Enter the location of the AutoCAD managed assemblies on your computer (these are in your ObjectARX SDK installation).
2. Click the 'Post-Process VB Express Project' button and select your vbproj file. (You can either delete your vbproj.user file or process it as well).

New features on this version:
1. Ability to work with different versions and platforms of AutoCAD
2. Enable AnyCPU compile options for Visual Basic Express (i.e. you can create .NET addins that also run on 64-bit operating systems)
3. Minor bug fixes

You are welcome to modify my code as you wish, but please include credit to Leslie Lowe with your version, and include his original project with it.

Please send me any enhancements you make to the project.

Augusto Goncalves
Autodesk Developer Network
augusto.goncalves@autodesk.com
February 17 2010

__________________________________________________________________________
      Supported controls:
        MSForms.Label
        MSForms.TextBox
        MSForms.CheckBox
        MSForms.OptionButton
        MSForms.CommandButton
        MSForms.ToggleButton
        MSForms.Image
        MSForms.ListBox
        MSForms.ComboBox
        MSForms.ScrollBar
        MSForms.Frame

      Unsupported Controls:
        MSForms.TabStrip
        MSForms.MultiPage
        MSForms.SpinButton
        All other external controls




