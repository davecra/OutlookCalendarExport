using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OutlookCalendarExport
{
    public class AddinSettings
    {
        public bool EmphasizeRecurring { get; set; }
        public bool DisplayTimeOnly { get; set; }
        public bool ShowLocation { get; set; }
        public ExportType ExportWhat {get; set; }
        public string PrintWhat { get; set; }
        public bool ShowHeader { get; set; }
        public ExtendedRecipientList Recipients { get; set; }
        public DateTime Date { get; set; }
        public enum TemplateType { Day, WorkWeek, FullWeek, Month };
        public enum ExportType {  All, Meetings, Shared };
        public bool ExcludePrivate { get; set; }
        /// <summary>
        /// Load the settings from the registry
        /// </summary>
        public void LoadSettings()
        {
            try
            {
                RegistryKey LobjKey = Registry.CurrentUser.OpenSubKey(Common.REGPATH, false);
                Recipients = new ExtendedRecipientList();
                string LstrList = LobjKey.GetValue("Recipients", "").ToString();
                Recipients.FromRegistryString(LstrList);
                if (Recipients.Count == 0)
                {
                    Recipients.Add(new ExtendedRecipient(Globals.ThisAddIn.Application.Session.CurrentUser));
                }
                this.PrintWhat = LobjKey.GetValue("PrintWhat", "").ToString();
                this.ShowHeader = bool.Parse(LobjKey.GetValue("ShowHeader", 1).ToString());
                this.ExportWhat = LobjKey.GetValue("ExportWhat", ExportType.All).ToString().GetEnumFromName<ExportType>();
                this.ShowLocation = bool.Parse(LobjKey.GetValue("ShowLocation", true).ToString());
                this.EmphasizeRecurring = bool.Parse(LobjKey.GetValue("EmphasizeRecurring", false).ToString());
                this.DisplayTimeOnly = bool.Parse(LobjKey.GetValue("DisplayTimeOnly", false).ToString());
                this.ExcludePrivate = bool.Parse(LobjKey.GetValue("ExcludePrivate", false).ToString());
            }
            catch(Exception PobjEx)
            {
                throw new Exception("Load settings from registry failed. " + PobjEx.Message);
            } 
        }

        /// <summary>
        /// Save the settings back to the registry
        /// </summary>
        public void SaveSettings()
        {
            // save settings
            try
            {
                RegistryKey LobjKey = Registry.CurrentUser.OpenSubKey(Common.REGPATH, true);
                if (LobjKey == null)
                {
                    LobjKey = Registry.CurrentUser.CreateSubKey(Common.REGPATH);
                }
                LobjKey.SetValue("Recipients", this.Recipients.ToRegistryString());
                LobjKey.SetValue("PrintWhat", this.PrintWhat);
                LobjKey.SetValue("ShowHeader", this.ShowHeader.ToString());
                LobjKey.SetValue("ExportWhat", this.ExportWhat.ToString());
                LobjKey.SetValue("ShowLocation", this.ShowLocation.ToString());
                LobjKey.SetValue("EmphasizeRecurring", this.EmphasizeRecurring.ToString());
                LobjKey.SetValue("DisplayTimeOnly", this.DisplayTimeOnly.ToString());
                LobjKey.SetValue("ExcludePrivate", this.ExcludePrivate.ToString());
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Failed to save settings to registry. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Returns a List of strings with recipient names.
        /// </summary>
        /// <returns></returns>
        public List<string> GetRecipientList()
        {
            try
            {
                List<string> LobjReturn = new List<string>();
                foreach (ExtendedRecipient LobjItem in Recipients)
                {
                    LobjReturn.Add(LobjItem.RecipientName);
                }
                return LobjReturn;
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log();
                return null;
            }
        }

        /// <summary>
        /// Determines the template type from the name
        /// </summary>
        /// <returns></returns>
        public TemplateType GetTemplateType()
        {
            if (PrintWhat.ToUpper().StartsWith("[DAY]"))
            {
                return TemplateType.Day;
            }
            else if (PrintWhat.ToUpper().StartsWith("[WORKWEEK]"))
            {
                return TemplateType.WorkWeek;
            }
            else if (PrintWhat.ToUpper().StartsWith("[FULLWEEK]"))
            {
                return TemplateType.FullWeek;
            }
            else
            {
                return TemplateType.Month;
            }
        }
    }
}
