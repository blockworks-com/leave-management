using LeaveManagement.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace LeaveManagement.OutlookAddIn2010
{
    public partial class ThisAddIn
    {
        #region Instance Variables

        public stdole.IPictureDisp _pictdisp = null;

        // List of tracked inspector windows
        internal static List<OutlookInspector> _inspectorWindows;

        // Ribbon UI reference
        internal static Office.IRibbonUI _ribbon;

        // List of tracked explorer windows
        internal static List<OutlookExplorer> _windows;

        private Outlook.Application _application;
        private Outlook.Explorers _explorers;
        private Outlook.MAPIFolder _inbox;
        private Outlook.Inspectors _inspectors;

        private Outlook.Items _items;

        // Inbox variables
        private Outlook.NameSpace _outlookNameSpace;

        #endregion Instance Variables

        #region VSTO Startup and Shutdown methods

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            DateTime startTime = DateTime.Now;
            LeaveManagementRibbonAddin result = null;

            // Use a generic try-catch to ensure no exceptions are thrown to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                result = new LeaveManagementRibbonAddin(_application);
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }

            return result;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            DateTime startTime = DateTime.Now;

            // Use a generic try-catch to ensure no exceptions are thrown to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                // Unhook event handlers
                _items.ItemAdd -=
                Inbox_NewItem;

                _explorers.NewExplorer -=
                    new Outlook.ExplorersEvents_NewExplorerEventHandler(
                        Explorers_NewExplorer);
                _inspectors.NewInspector -=
                    new Outlook.InspectorsEvents_NewInspectorEventHandler(
                        Inspectors_NewInspector);

                // Dereference objects
                _items = null;
                _inbox = null;
                _outlookNameSpace = null;
                _pictdisp = null;
                _explorers = null;
                _inspectors = null;
                _windows.Clear();
                _windows = null;
                _inspectorWindows.Clear();
                _inspectorWindows = null;
                _ribbon = null;
                _application = null;
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            DateTime startTime = DateTime.Now;

            // Use a generic try-catch to ensure no uncaught exceptions go to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                // Initialize variables
                _application = this.Application;
                _explorers = _application.Explorers;
                _inspectors = _application.Inspectors;
                _windows = new List<OutlookExplorer>();
                _inspectorWindows = new List<OutlookInspector>();

                _outlookNameSpace = Application.GetNamespace("MAPI");
                _inbox = _outlookNameSpace.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox);

                // Wire up event handlers to handle multiple Explorer windows
                _explorers.NewExplorer +=
                    new Outlook.ExplorersEvents_NewExplorerEventHandler(
                        Explorers_NewExplorer);

                // Wire up event handlers to handle multiple Inspector windows
                _inspectors.NewInspector +=
                    new Outlook.InspectorsEvents_NewInspectorEventHandler(
                        Inspectors_NewInspector);

                // Add the ActiveExplorer to _windows
                Outlook.Explorer expl = _application.ActiveExplorer()
                    as Outlook.Explorer;
                OutlookExplorer window = new OutlookExplorer(expl);
                _windows.Add(window);

                // Hook up event handlers for window
                window.Close += new EventHandler(WrappedWindow_Close);
                window.InvalidateControl += new EventHandler<
                    OutlookExplorer.InvalidateEventArgs>(
                    WrappedWindow_InvalidateControl);

                // Wire up event handler to handle incoming inbox items
                _items = _inbox.Items;
                _items.ItemAdd +=
                    Inbox_NewItem;

                // Get IPictureDisp for CurrentUser on startup
                try
                {
                    Outlook.AddressEntry addrEntry =
                        Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
                    if (addrEntry.Type == "EX")
                    {
                        Outlook.ExchangeUser exchUser =
                            addrEntry.GetExchangeUser() as Outlook.ExchangeUser;
                        _pictdisp = exchUser.GetPicture() as stdole.IPictureDisp;
                    }
                }
                catch (Exception ex)
                {
                    LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
                }
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }
        }

        // Note: Outlook no longer raises the Shutdown event. If you have code that must run when Outlook shuts down,
        //       see http://go.microsoft.com/fwlink/?LinkId=506785

        #endregion VSTO Startup and Shutdown methods

        #region Helper Methods

        /// <summary>
        /// Looks up the window wrapper for a given window object
        /// </summary>
        /// <param name="window">An outlook explorer window</param>
        /// <returns></returns>
        internal static OutlookExplorer FindOutlookExplorer(object window)
        {
            foreach (OutlookExplorer Explorer in _windows)
            {
                if (Explorer.Window == window)
                {
                    return Explorer;
                }
            }
            return null;
        }

        /// <summary>
        /// Looks up the window wrapper for a given window object
        /// </summary>
        /// <param name="window">An outlook inspector window</param>
        /// <returns></returns>
        internal static OutlookInspector FindOutlookInspector(object window)
        {
            foreach (OutlookInspector Inspector in _inspectorWindows)
            {
                if (Inspector.Window == window)
                {
                    return Inspector;
                }
            }
            return null;
        }

        #endregion Helper Methods

        #region Event Handers

        /// <summary>
        /// The NewExplorer event fires whenever a new Explorer is displayed.
        /// </summary>
        /// <param name="Explorer"></param>
        private void Explorers_NewExplorer(Outlook.Explorer Explorer)
        {
            DateTime startTime = DateTime.Now;

            // Use a generic try-catch to ensure no exceptions are thrown to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                // Check to see if this is a new window we don't already track
                OutlookExplorer existingWindow =
                    FindOutlookExplorer(Explorer);
                // If the _windows collection does not have a window for this Explorer, we should add it to m_Windows
                if (existingWindow == null)
                {
                    OutlookExplorer window = new OutlookExplorer(Explorer);
                    if (window != null)
                    {
                        window.Close += new EventHandler(WrappedWindow_Close);
                        window.InvalidateControl += new EventHandler<
                            OutlookExplorer.InvalidateEventArgs>(
                            WrappedWindow_InvalidateControl);
                        _windows.Add(window);
                    }
                }
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }
        }

        /// <summary>
        /// The NewItem event fires whenever a new Item arrives in the inbox.
        /// </summary>
        /// <param name="Item"></param>
        private void Inbox_NewItem(object Item)
        {
            DateTime startTime = DateTime.Now;
            string filter = "New Hire Advisement - ";

            // Use a generic try-catch to ensure no exceptions are thrown to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                OutlookItem olItem = new OutlookItem(Item);
                if (olItem != null)
                {
                    if (olItem.MessageClass == "IPM.Note" &&
                        olItem.Subject.StartsWith(filter, true, null))
                    {
                        string msg = "New Joiner item in Inbox";
                        //MessageBox.Show(msg, "LeaveManagement.OutlookAddIn2010",
                        //    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }
        }

        /// <summary>
        /// The NewInspector event fires whenever a new Inspector is displayed.
        /// </summary>
        /// <param name="Explorer"></param>
        private void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            DateTime startTime = DateTime.Now;

            // Use a generic try-catch to ensure no exceptions are thrown to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                if (_ribbon != null)
                {
                    _ribbon.Invalidate();
                }

                // Check to see if this is a new window we don't already track
                OutlookInspector existingWindow =
                    FindOutlookInspector(Inspector);
                // If the m_InspectorWindows collection does not have a window for this Inspector, we should add it to m_InspectorWindows
                if (existingWindow == null)
                {
                    OutlookInspector window = new OutlookInspector(Inspector);
                    if (window != null)
                    {
                        window.Close += new EventHandler(WrappedInspectorWindow_Close);
                        window.InvalidateControl += new EventHandler<
                            OutlookInspector.InvalidateEventArgs>(
                            WrappedInspectorWindow_InvalidateControl);
                        _inspectorWindows.Add(window);
                    }
                }
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }
        }

        private void WrappedInspectorWindow_Close(object sender, EventArgs e)
        {
            DateTime startTime = DateTime.Now;

            // Use a generic try-catch to ensure no exceptions are thrown to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                OutlookInspector window = (OutlookInspector)sender;
                if (window != null)
                {
                    window.Close -= new EventHandler(WrappedInspectorWindow_Close);
                    _inspectorWindows.Remove(window);
                }
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }
        }

        private void WrappedInspectorWindow_InvalidateControl(object sender,
                    OutlookInspector.InvalidateEventArgs e)
        {
            DateTime startTime = DateTime.Now;

            // Use a generic try-catch to ensure no exceptions are thrown to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                if (_ribbon != null)
                {
                    _ribbon.InvalidateControl(e.ControlID);
                }
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }
        }

        private void WrappedWindow_Close(object sender, EventArgs e)
        {
            DateTime startTime = DateTime.Now;

            // Use a generic try-catch to ensure no exceptions are thrown to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                OutlookExplorer window = (OutlookExplorer)sender;
                if (window != null)
                {
                    window.Close -= new EventHandler(WrappedWindow_Close);
                    _windows.Remove(window);
                }
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }
        }

        private void WrappedWindow_InvalidateControl(object sender,
                    OutlookExplorer.InvalidateEventArgs e)
        {
            DateTime startTime = DateTime.Now;

            // Use a generic try-catch to ensure no exceptions are thrown to Outlook.
            try
            {
                LogWrapper.MainLogger.Debug(string.Format("Entering method '{0}'", MethodBase.GetCurrentMethod().Name));

                if (_ribbon != null)
                {
                    _ribbon.InvalidateControl(e.ControlID);
                }
            }
            catch (Exception ex)
            {
                LogWrapper.MainLogger.Error(ex, string.Format("Exception in method '{0}'", MethodBase.GetCurrentMethod().Name));
            }
            finally
            {
                LogWrapper.MainLogger.Debug(string.Format("Exiting method '{0}' took '{1}' milliseconds", MethodBase.GetCurrentMethod().Name, ((TimeSpan)(DateTime.Now - startTime)).Milliseconds));
            }
        }

        #endregion Event Handers

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion VSTO generated code
    }
}