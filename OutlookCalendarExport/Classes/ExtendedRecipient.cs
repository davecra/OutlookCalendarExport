using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarExport
{
    /// <summary>
    /// Recipient Display Option Class
    /// This class helps us keep track of recipient names to display symbols and colors
    /// </summary>
    public class ExtendedRecipient : IEquatable<ExtendedRecipient>
    {
        [XmlIgnore]
        public Outlook.Recipient InteropRecipient { get; private set; }
        public bool ShowName { get; set; }
        public string RecipientName { get; set; }
        public string DisplayName { get; set; }
        public string Symbol { get; set; }
        public string HighlightColor { get; set; }
        public string EntryId { get; set; }
        [XmlIgnore]
        public List<ExtendedAppointment> Appointments { get; private set; }

        /// <summary>
        /// ctor - parameterless for serialization
        /// </summary>
        public ExtendedRecipient() { }

        /// <summary>
        /// Creates a new recipient with extended properties
        /// </summary>
        /// <param name="PstrName"></param>
        /// <param name="PbolShow"></param>
        /// <param name="PstrDisplayName"></param>
        /// <param name="PstrColor"></param>
        /// <param name="PstrSymbol"></param>
        public ExtendedRecipient(string PstrName, string PstrId, bool PbolShow, string PstrDisplayName, Color PstrColor, string PstrSymbol)
        {
            try
            {
                RecipientName = PstrName;
                ShowName = PbolShow;
                DisplayName = PstrDisplayName;
                HighlightColor = PstrColor.ToRGBColorString();
                Symbol = PstrSymbol;
                EntryId = PstrId;
                InteropRecipient = Globals.ThisAddIn.Application.Session.GetRecipientFromID(PstrId);
                if (!InteropRecipient.Resolve())
                {
                    throw new Exception("Unable to resolve " + PstrName + ".");
                }
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Creating extended recipient failed. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Creates a new recipient with extended properties
        /// </summary>
        /// <param name="PobjRecipient"></param>
        public ExtendedRecipient(Outlook.Recipient PobjRecipient)
        {
            try
            {
                RecipientName = PobjRecipient.Name;
                ShowName = true;
                DisplayName = PobjRecipient.Name;
                HighlightColor = Color.Black.ToRGBColorString();
                Symbol = ""; // none
                InteropRecipient = PobjRecipient;
                EntryId = PobjRecipient.EntryID;
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Creating extended recipient failed. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Used for .Contains for easy List operation
        /// </summary>
        /// <param name="PobjOther"></param>
        /// <returns></returns>
        public bool Equals(ExtendedRecipient PobjOther)
        {
            return PobjOther.RecipientName == RecipientName;
        }
    }
}
