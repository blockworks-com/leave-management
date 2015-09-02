using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace LeaveManagement.OutlookAddIn2010
{
    /// <summary>
    /// This class tracks the state of an Outlook Inspector window for your add-in and ensures that what happens in this
    /// window is handled correctly.
    /// </summary>
    internal class OutlookInspector
    {
        #region Instance Variables

        private Outlook.AppointmentItem _appointment;

        // wrapped AppointmentItem
        private Outlook.ContactItem _contact;

        // Use these instance variables to handle item-level events
        private Outlook.MailItem _mail;

        // wrapped ContactItem
        private Outlook.ContactItem _task;

        private Outlook.Inspector _window;             // wrapped window object

        // wrapped MailItem

        // wrapped TaskItem Define other class-level item instance variables as needed

        #endregion Instance Variables

        #region Events

        public event EventHandler Close;

        public event EventHandler<InvalidateEventArgs> InvalidateControl;

        #endregion Events

        #region Constructor

        /// <summary>
        /// Create a new instance of the tracking class for a particular
        /// inspector and custom task pane.
        /// </summary>
        /// <param name="inspector">A new inspector window to track</param>
        ///<remarks></remarks>
        public OutlookInspector(Outlook.Inspector inspector)
        {
            _window = inspector;

            // Hookup the close event
            ((Outlook.InspectorEvents_Event)inspector).Close +=
                new Outlook.InspectorEvents_CloseEventHandler(
                OutlookInspectorWindow_Close);

            // Hookup item-level events as needed
            // For example, the following code hooks up PropertyChange
            // event for a ContactItem
            //OutlookItem olItem = new OutlookItem(inspector.CurrentItem);
            //if(olItem.Class==Outlook.OlObjectClass.olContact)
            //{
            //    m_Contact = olItem.InnerObject as Outlook.ContactItem;
            //    m_Contact.PropertyChange +=
            //        new Outlook.ItemEvents_10_PropertyChangeEventHandler(
            //        m_Contact_PropertyChange);
            //}
        }

        #endregion Constructor

        #region Event Handlers

        /// <summary>
        /// Event Handler for the inspector close event.
        /// </summary>
        private void OutlookInspectorWindow_Close()
        {
            // Unhook events from any item-level instance variables
            //m_Contact.PropertyChange -=
            //    Outlook.ItemEvents_10_PropertyChangeEventHandler(
            //    m_Contact_PropertyChange);

            // Unhook events from the window
            ((Outlook.InspectorEvents_Event)_window).Close -=
                new Outlook.InspectorEvents_CloseEventHandler(
                OutlookInspectorWindow_Close);

            // Raise the OutlookInspector close event
            if (Close != null)
            {
                Close(this, EventArgs.Empty);
            }

            // Unhook any item-level instance variables
            //m_Contact = null;
            _window = null;
        }

        //void  m_Contact_PropertyChange(string Name)
        //{
        //    // Implement PropertyChange here
        //}

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
        /// The actual Outlook inspector window wrapped by this instance
        /// </summary>
        internal Outlook.Inspector Window
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