using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace LeaveManagement.OutlookAddIn2010
{
    /// <summary>
    /// This class tracks the state of an Outlook Explorer window for your add-in and ensures that what happens in this
    /// window is handled correctly.
    /// </summary>
    internal class OutlookExplorer
    {
        #region Instance Variables

        private Outlook.Explorer _window;   // wrapped window object

        #endregion Instance Variables

        #region Events

        public event EventHandler Close;

        public event EventHandler<InvalidateEventArgs> InvalidateControl;

        #endregion Events

        #region Constructor

        /// <summary>
        /// Create a new instance of the tracking class for a particular explorer
        /// </summary>
        /// <param name="explorer">A new explorer window to track</param>
        ///<remarks></remarks>
        public OutlookExplorer(Outlook.Explorer explorer)
        {
            _window = explorer;

            // Hookup Close event
            ((Outlook.ExplorerEvents_Event)explorer).Close +=
                new Outlook.ExplorerEvents_CloseEventHandler(
                OutlookExplorerWindow_Close);

            // Hookup SelectionChange event
            _window.SelectionChange +=
                new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(
                    Window_SelectionChange);
        }

        #endregion Constructor

        #region Event Handlers

        /// <summary>
        /// Event Handler for Close event.
        /// </summary>
        private void OutlookExplorerWindow_Close()
        {
            // Unhook explorer-level events

            _window.SelectionChange -=
                new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(
                Window_SelectionChange);

            ((Outlook.ExplorerEvents_Event)_window).Close -=
                new Outlook.ExplorerEvents_CloseEventHandler(
                OutlookExplorerWindow_Close);

            // Raise the OutlookExplorer close event
            if (Close != null)
            {
                Close(this, EventArgs.Empty);
            }

            _window = null;
        }

        /// <summary>
        /// Event Handler for SelectionChange event
        /// </summary>
        private void Window_SelectionChange()
        {
            RaiseInvalidateControl("MyTab");
        }

        #endregion Event Handlers

        #region Methods

        private void RaiseInvalidateControl(string controlID)
        {
            if (InvalidateControl != null)
                InvalidateControl(this, new InvalidateEventArgs(controlID));
        }

        #endregion Methods

        #region Properties

        /// <summary>
        /// The actual Outlook explorer window wrapped by this instance
        /// </summary>
        internal Outlook.Explorer Window
        {
            get { return _window; }
        }

        #endregion Properties

        #region Helper Class

        public class InvalidateEventArgs : EventArgs
        {
            private string _controlID;

            public InvalidateEventArgs(string controlID)
            {
                _controlID = controlID;
            }

            public string ControlID
            {
                get { return _controlID; }
            }
        }

        #endregion Helper Class
    }
}