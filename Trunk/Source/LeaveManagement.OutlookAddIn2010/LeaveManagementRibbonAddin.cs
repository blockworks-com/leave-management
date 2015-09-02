﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace LeaveManagement.OutlookAddIn2010
{
    [ComVisible(true)]
    public class LeaveManagementRibbonAddin : Office.IRibbonExtensibility
    {
        private Outlook.Application _outlookApplication;
        private Office.IRibbonUI _ribbon;

        //Override of constructor to pass
        // a trusted Outlook.Application object
        public LeaveManagementRibbonAddin(Outlook.Application outlookApplication)
        {
            _outlookApplication = outlookApplication as Outlook.Application;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string customUI = string.Empty;
            Debug.WriteLine(ribbonID);

            //Return the appropriate Ribbon XML for ribbonID
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                    customUI = GetResourceText(
                        "LeaveManagement.OutlookAddIn2010.Explorer.xml");
                    return customUI;

                case "Microsoft.Outlook.Mail.Read":
                    customUI = GetResourceText(
                        "LeaveManagement.OutlookAddIn2010.ReadMail.xml");
                    return customUI;

                default:
                    return string.Empty;
            }
        }

        #endregion IRibbonExtensibility Members

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ThisAddIn._ribbon = ribbonUI;
        }

        #region Visibility Callbacks

        // Hide Backstage in an Inspector window
        public bool Backstage_OnGetVisible(Office.IRibbonControl control)
        {
            if (control.Context is Outlook.Explorer)
                return true;
            else
                return false;
        }

        // Only show Buttons if appropriate
        public bool OnGetVisible(Office.IRibbonControl control)
        {
            //Contract.Requires(null != control);

            if (null == control)
            {
                return false;
            }

            bool result = false;
            switch (control.Id)
            {
                case "LmNewHireButton":
                    result = true;
                    break;

                case "LmAdjustmentButton":
                    result = true;
                    break;

                case "LmLeaverButton":
                    result = true;
                    break;

                case "LmDelegateButton":
                    result = true;
                    break;

                case "LmPendingButton":
                    result = true;
                    break;

                case "LmLeaveRequestButton":
                    result = false;
                    break;
            }

            return result;
        }

        // Only show Tab when Explorer Selection is a received mail or when Inspector is a read note
        public bool Tab_OnGetVisible(Office.IRibbonControl control)
        {
            if (control.Context is Outlook.Explorer)
            {
                Outlook.Explorer explorer =
                    control.Context as Outlook.Explorer;
                Outlook.Selection selection = explorer.Selection;
                if (selection.Count == 1)
                {
                    if (selection[1] is Outlook.MailItem)
                    {
                        Outlook.MailItem oMail =
                            selection[1] as Outlook.MailItem;
                        if (oMail.Sent == true)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else if (control.Context is Outlook.Inspector)
            {
                Outlook.Inspector oInsp =
                    control.Context as Outlook.Inspector;
                if (oInsp.CurrentItem is Outlook.MailItem)
                {
                    Outlook.MailItem oMail =
                        oInsp.CurrentItem as Outlook.MailItem;
                    if (oMail.Sent == true)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        public bool TabInspector_OnGetVisible(Office.IRibbonControl control)
        {
            if (control.Context is Outlook.Inspector)
            {
                Outlook.Inspector oInsp =
                    control.Context as Outlook.Inspector;
                if (oInsp.CurrentItem is Outlook.MailItem)
                {
                    Outlook.MailItem oMail =
                        oInsp.CurrentItem as Outlook.MailItem;
                    if (oMail.Sent == true)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return true;
            }
        }

        #endregion Visibility Callbacks

        #region Click Callbacks

        public void OnAdjustmentButtonClick(Office.IRibbonControl control)
        {
            List<OutlookItem> items = GetItems(control);
            string msg = "Adjustment";
            MessageBox.Show(msg, "LeaveManagement.OutlookAddIn2010",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnDelegateButtonClick(Office.IRibbonControl control)
        {
            List<OutlookItem> items = GetItems(control);
            string msg = "Delegate";
            MessageBox.Show(msg, "LeaveManagement.OutlookAddIn2010",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnLeaverButtonClick(Office.IRibbonControl control)
        {
            List<OutlookItem> items = GetItems(control);
            string msg = "Leaver";
            MessageBox.Show(msg, "LeaveManagement.OutlookAddIn2010",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnLeaveRequestButtonClick(Office.IRibbonControl control)
        {
            List<OutlookItem> items = GetItems(control);
            string msg = "Leave request";
            MessageBox.Show(msg, "LeaveManagement.OutlookAddIn2010",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnNewHireButtonClick(Office.IRibbonControl control)
        {
            List<OutlookItem> items = GetItems(control);
            string msg = "Joiner";
            MessageBox.Show(msg, "LeaveManagement.OutlookAddIn2010",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnPendingButtonClick(Office.IRibbonControl control)
        {
            string msg = "Pending list";
            MessageBox.Show(msg, "LeaveManagement.OutlookAddIn2010",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion Click Callbacks

        #region Helpers

        public stdole.IPictureDisp GetCurrentUserImage(Office.IRibbonControl control)
        {
            //stdole.IPictureDisp pictureDisp = null;
            Outlook.AddressEntry addrEntry =
                Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                if (Globals.ThisAddIn._pictdisp != null)
                {
                    return Globals.ThisAddIn._pictdisp;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private List<OutlookItem> GetItems(Office.IRibbonControl control)
        {
            List<OutlookItem> result = new List<OutlookItem>();

            string msg = string.Empty;
            if (control.Context is Outlook.AttachmentSelection)
            {
                Outlook.AttachmentSelection attachSel =
                    control.Context as Outlook.AttachmentSelection;
                foreach (Outlook.Attachment attach in attachSel)
                {
                }
            }
            else if (control.Context is Outlook.Folder)
            {
                Outlook.Folder folder =
                    control.Context as Outlook.Folder;
            }
            else if (control.Context is Outlook.Selection)
            {
                Outlook.Selection selection =
                    control.Context as Outlook.Selection;
                for (int i = 0; i < selection.Count; i++)
                {
                    OutlookItem olItem =
                        new OutlookItem(selection[i + 1]); // 1 based index
                    result.Add(olItem);
                }
            }
            else if (control.Context is Outlook.OutlookBarShortcut)
            {
                Outlook.OutlookBarShortcut shortcut =
                    control.Context as Outlook.OutlookBarShortcut;
            }
            else if (control.Context is Outlook.Store)
            {
                Outlook.Store store =
                    control.Context as Outlook.Store;
            }
            else if (control.Context is Outlook.View)
            {
                Outlook.View view =
                    control.Context as Outlook.View;
            }
            else if (control.Context is Outlook.Inspector)
            {
                Outlook.Inspector insp =
                    control.Context as Outlook.Inspector;
                OutlookItem olItem =
                    new OutlookItem(insp.CurrentItem);
                result.Add(olItem);
            }
            else if (control.Context is Outlook.Explorer)
            {
                Outlook.Explorer explorer =
                    control.Context as Outlook.Explorer;
                Outlook.Selection selection =
                    explorer.Selection;
                for (int i = 0; i < selection.Count; i++)
                {
                    OutlookItem olItem =
                        new OutlookItem(selection[i + 1]); // 1 based index
                    result.Add(olItem);
                }
            }
            else if (control.Context is Outlook.NavigationGroup)
            {
                Outlook.NavigationGroup navGroup =
                    control.Context as Outlook.NavigationGroup;
            }
            else if (control.Context is
                Microsoft.Office.Core.IMsoContactCard)
            {
                Office.IMsoContactCard card =
                    control.Context as Office.IMsoContactCard;
                if (card.AddressType ==
                    Office.MsoContactCardAddressType.
                    msoContactCardAddressTypeOutlook)
                {
                    // IMSOContactCard.Address is AddressEntry.ID
                    Outlook.AddressEntry addr =
                        Globals.ThisAddIn.Application.Session.GetAddressEntryFromID(
                        card.Address);
                    if (addr != null)
                    {
                    }
                }
            }
            else if (control.Context is Outlook.NavigationModule)
            {
            }
            else if (control.Context == null)
            {
            }
            else
            {
            }

            return result;
        }

        #endregion Helpers
    }
}