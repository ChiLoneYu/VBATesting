using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Vbe.Interop;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace VBEAddin
{
    [ComVisible(true), Guid("3599862B-FF92-42DF-BB55-DBD37CC13565"), ProgId("VBEAddIn.Connect")]
    public class Connect : IDTExtensibility2
    {
        private VBE _VBE;
        private AddIn _AddIn;

        // Buttons created by the add-in
        private CommandBarButton withEventsField__myStandardCommandBarButton;
        private CommandBarButton _myStandardCommandBarButton
        {
            get { return withEventsField__myStandardCommandBarButton; }
            set
            {
                if (withEventsField__myStandardCommandBarButton != null)
                {
                    withEventsField__myStandardCommandBarButton.Click -= _myStandardCommandBarButton_Click;
                }
                withEventsField__myStandardCommandBarButton = value;
                if (withEventsField__myStandardCommandBarButton != null)
                {
                    withEventsField__myStandardCommandBarButton.Click += _myStandardCommandBarButton_Click;
                }
            }
        }
        private CommandBarButton withEventsField__myToolsCommandBarButton;
        private CommandBarButton _myToolsCommandBarButton
        {
            get { return withEventsField__myToolsCommandBarButton; }
            set
            {
                if (withEventsField__myToolsCommandBarButton != null)
                {
                    withEventsField__myToolsCommandBarButton.Click -= _myToolsCommandBarButton_Click;
                }
                withEventsField__myToolsCommandBarButton = value;
                if (withEventsField__myToolsCommandBarButton != null)
                {
                    withEventsField__myToolsCommandBarButton.Click += _myToolsCommandBarButton_Click;
                }
            }
        }
        private CommandBarButton withEventsField__myCodeWindowCommandBarButton;
        private CommandBarButton _myCodeWindowCommandBarButton
        {
            get { return withEventsField__myCodeWindowCommandBarButton; }
            set
            {
                if (withEventsField__myCodeWindowCommandBarButton != null)
                {
                    withEventsField__myCodeWindowCommandBarButton.Click -= _myCodeWindowCommandBarButton_Click;
                }
                withEventsField__myCodeWindowCommandBarButton = value;
                if (withEventsField__myCodeWindowCommandBarButton != null)
                {
                    withEventsField__myCodeWindowCommandBarButton.Click += _myCodeWindowCommandBarButton_Click;
                }
            }
        }
        private CommandBarButton withEventsField__myToolBarButton;
        private CommandBarButton _myToolBarButton
        {
            get { return withEventsField__myToolBarButton; }
            set
            {
                if (withEventsField__myToolBarButton != null)
                {
                    withEventsField__myToolBarButton.Click -= _myToolBarButton_Click;
                }
                withEventsField__myToolBarButton = value;
                if (withEventsField__myToolBarButton != null)
                {
                    withEventsField__myToolBarButton.Click += _myToolBarButton_Click;
                }
            }
        }
        private CommandBarButton withEventsField__myCommandBarPopup1Button;
        private CommandBarButton _myCommandBarPopup1Button
        {
            get { return withEventsField__myCommandBarPopup1Button; }
            set
            {
                if (withEventsField__myCommandBarPopup1Button != null)
                {
                    withEventsField__myCommandBarPopup1Button.Click -= _myCommandBarPopup1Button_Click;
                }
                withEventsField__myCommandBarPopup1Button = value;
                if (withEventsField__myCommandBarPopup1Button != null)
                {
                    withEventsField__myCommandBarPopup1Button.Click += _myCommandBarPopup1Button_Click;
                }
            }
        }
        private CommandBarButton withEventsField__myCommandBarPopup2Button;
        private CommandBarButton _myCommandBarPopup2Button
        {
            get { return withEventsField__myCommandBarPopup2Button; }
            set
            {
                if (withEventsField__myCommandBarPopup2Button != null)
                {
                    withEventsField__myCommandBarPopup2Button.Click -= _myCommandBarPopup2Button_Click;
                }
                withEventsField__myCommandBarPopup2Button = value;
                if (withEventsField__myCommandBarPopup2Button != null)
                {
                    withEventsField__myCommandBarPopup2Button.Click += _myCommandBarPopup2Button_Click;
                }
            }

        }
        // CommandBars created by the add-in
        private CommandBar _myToolbar;
        private CommandBarPopup _myCommandBarPopup1;

        private CommandBarPopup _myCommandBarPopup2;

        #region "IDTExtensibility2 Members"

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _VBE = (VBE)application;
                _AddIn = (AddIn)addInInst;

                switch (connectMode)
                {
                    case Extensibility.ext_ConnectMode.ext_cm_Startup:
                        break;
                    case Extensibility.ext_ConnectMode.ext_cm_AfterStartup:
                        InitializeAddIn();

                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void onReferenceItemAdded(Reference reference)
        {
            //TODO: Map types found in assembly using reference.
        }

        private void onReferenceItemRemoved(Reference reference)
        {
            //TODO: Remove types found in assembly using reference.
        }

        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
            try
            {
                switch (disconnectMode)
                {

                    case ext_DisconnectMode.ext_dm_HostShutdown:
                    case ext_DisconnectMode.ext_dm_UserClosed:

                        // Delete buttons on built-in commandbars
                        if ((_myStandardCommandBarButton != null))
                        {
                            _myStandardCommandBarButton.Delete();
                        }

                        if ((_myCodeWindowCommandBarButton != null))
                        {
                            _myCodeWindowCommandBarButton.Delete();
                        }

                        if ((_myToolsCommandBarButton != null))
                        {
                            _myToolsCommandBarButton.Delete();
                        }

                        // Disconnect event handlers
                        _myToolBarButton = null;
                        _myCommandBarPopup1Button = null;
                        _myCommandBarPopup2Button = null;

                        // Delete commandbars created by the add-in
                        if ((_myToolbar != null))
                        {
                            _myToolbar.Delete();
                        }

                        if ((_myCommandBarPopup1 != null))
                        {
                            _myCommandBarPopup1.Delete();
                        }

                        if ((_myCommandBarPopup2 != null))
                        {
                            _myCommandBarPopup2.Delete();
                        }

                        break;
                }

            }
            catch (System.Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
            InitializeAddIn();
        }

        private CommandBarButton AddCommandBarButton(CommandBar commandBar)
        {

            CommandBarButton commandBarButton = default(CommandBarButton);
            CommandBarControl commandBarControl = default(CommandBarControl);

            commandBarControl = commandBar.Controls.Add(MsoControlType.msoControlButton);
            commandBarButton = (CommandBarButton)commandBarControl;

            commandBarButton.Caption = "My button";
            commandBarButton.FaceId = 59;

            return commandBarButton;

        }

        private void InitializeAddIn()
        {
            //MessageBox.Show(_AddIn.ProgId + " loaded in VBA editor version " + _VBE.Version);
            // Constants for names of built-in commandbars of the VBA editor
            const string STANDARD_COMMANDBAR_NAME = "Standard";
            const string MENUBAR_COMMANDBAR_NAME = "Menu Bar";
            const string TOOLS_COMMANDBAR_NAME = "Tools";
            const string CODE_WINDOW_COMMANDBAR_NAME = "Code Window";

            // Constants for names of commandbars created by the add-in
            const string MY_COMMANDBAR_POPUP1_NAME = "MyTemporaryCommandBarPopup1";
            const string MY_COMMANDBAR_POPUP2_NAME = "MyTemporaryCommandBarPopup2";

            // Constants for captions of commandbars created by the add-in
            const string MY_COMMANDBAR_POPUP1_CAPTION = "My sub menu";
            const string MY_COMMANDBAR_POPUP2_CAPTION = "My main menu";
            const string MY_TOOLBAR_CAPTION = "My toolbar";

            // Built-in commandbars of the VBA editor
            CommandBar standardCommandBar = default(CommandBar);
            CommandBar menuCommandBar = default(CommandBar);
            CommandBar toolsCommandBar = default(CommandBar);
            CommandBar codeCommandBar = default(CommandBar);

            // Other variables
            CommandBarControl toolsCommandBarControl = default(CommandBarControl);
            int position = 0;


            try
            {
                // Retrieve some built-in commandbars
                standardCommandBar = _VBE.CommandBars[STANDARD_COMMANDBAR_NAME];
                menuCommandBar = _VBE.CommandBars[MENUBAR_COMMANDBAR_NAME];
                toolsCommandBar = _VBE.CommandBars[TOOLS_COMMANDBAR_NAME];
                codeCommandBar = _VBE.CommandBars[CODE_WINDOW_COMMANDBAR_NAME];

               

                // Add a button to the built-in "Standard" toolbar
                _myStandardCommandBarButton = AddCommandBarButton(standardCommandBar);

                // Add a button to the built-in "Tools" menu
                _myToolsCommandBarButton = AddCommandBarButton(toolsCommandBar);

                // Add a button to the built-in "Code Window" context menu
                _myCodeWindowCommandBarButton = AddCommandBarButton(codeCommandBar);

                // ------------------------------------------------------------------------------------
                // New toolbar
                // ------------------------------------------------------------------------------------

                // Add a new toolbar 
                _myToolbar = _VBE.CommandBars.Add(MY_TOOLBAR_CAPTION, MsoBarPosition.msoBarTop, System.Type.Missing, true);
                
                // Add a new button on that toolbar
                _myToolBarButton = AddCommandBarButton(_myToolbar);

                // Make visible the toolbar
                _myToolbar.Visible = true;

                // ------------------------------------------------------------------------------------
                // New submenu under the "Tools" menu
                // ------------------------------------------------------------------------------------

                // Add a new commandbar popup 
                _myCommandBarPopup1 = (CommandBarPopup)toolsCommandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, toolsCommandBar.Controls.Count + 1, true);

                // Change some commandbar popup properties
                _myCommandBarPopup1.CommandBar.Name = MY_COMMANDBAR_POPUP1_NAME;
                _myCommandBarPopup1.Caption = MY_COMMANDBAR_POPUP1_CAPTION;

                // Add a new button on that commandbar popup
                _myCommandBarPopup1Button = AddCommandBarButton(_myCommandBarPopup1.CommandBar);

                // Make visible the commandbar popup
                _myCommandBarPopup1.Visible = true;

                // ------------------------------------------------------------------------------------
                // New main menu
                // ------------------------------------------------------------------------------------

                // Calculate the position of a new commandbar popup to the right of the "Tools" menu
                toolsCommandBarControl = (CommandBarControl)toolsCommandBar.Parent;
                position = toolsCommandBarControl.Index + 1;

                // Add a new commandbar popup 
                _myCommandBarPopup2 = (CommandBarPopup)menuCommandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, position, true);

                // Change some commandbar popup properties
                _myCommandBarPopup2.CommandBar.Name = MY_COMMANDBAR_POPUP2_NAME;
                _myCommandBarPopup2.Caption = MY_COMMANDBAR_POPUP2_CAPTION;

                // Add a new button on that commandbar popup
                _myCommandBarPopup2Button = AddCommandBarButton(_myCommandBarPopup2.CommandBar);

                // Make visible the commandbar popup
                _myCommandBarPopup2.Visible = true;

            }
            catch (System.Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }

        private void _myToolBarButton_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("Clicked " + Ctrl.Caption);
            
        }


        private void _myToolsCommandBarButton_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("Clicked " + Ctrl.Caption);

        }


        private void _myStandardCommandBarButton_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("Clicked " + Ctrl.Caption);

        }


        private void _myCodeWindowCommandBarButton_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("Clicked " + Ctrl.Caption);

        }


        private void _myCommandBarPopup1Button_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("Clicked " + Ctrl.Caption);

        }


        private void _myCommandBarPopup2Button_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("Clicked " + Ctrl.Caption);
            MessageBox.Show(_VBE.ActiveVBProject.Name); 
            


        }
        public void OnBeginShutdown(ref Array custom)
        {
        }

        #endregion
    }
}
