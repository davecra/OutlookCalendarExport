using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace OutlookCalendarExport
{
    public class ExtendedRecipientList : List<ExtendedRecipient>
    {

        /// <summary>
        /// ctor - parameterless for serialization
        /// </summary>
        public ExtendedRecipientList() { } 

        /// <summary>
        /// Returns an instance of item based on an instance of another item that
        /// is in the provided colleciton of item
        /// </summary>
        /// <typeparam name="ExtendedAppointment"></typeparam>
        /// <param name="PobjCollection"></param>
        /// <param name="PobjPredicate"></param>
        /// <returns></returns>
        public ExtendedRecipient FindItem(ExtendedRecipient PobjOther)
        {
            try
            {
                foreach (ExtendedRecipient LobjItem in this)
                {
                    if (LobjItem.RecipientName == PobjOther.RecipientName) return LobjItem;
                }

                return default(ExtendedRecipient);
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Failed while comparing recipient: " + PobjOther.RecipientName + ". " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Takes this list of recipients and returns a string formatted for the registry
        /// </summary>
        /// <returns></returns>
        public string ToRegistryString()
        {
            try
            {
                string LstrRetVal = "";
                foreach (ExtendedRecipient LobjItem in this)
                {
                    LstrRetVal += LobjItem.DisplayName + ";" +
                                  LobjItem.HighlightColor + ";" +
                                  LobjItem.RecipientName + ";" +
                                  LobjItem.EntryId + ";" +
                                  LobjItem.ShowName.ToString() + ";" +
                                  LobjItem.Symbol + "|";
                }
                // trim tailing |
                LstrRetVal = LstrRetVal.Substring(0, LstrRetVal.Length - 1);
                return LstrRetVal;
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Failed to write recipient list ot the registry. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Reads the values from a string that was stored in the registry
        /// and converts it to a list of recipients
        /// </summary>
        /// <param name="PobjValue"></param>
        public void FromRegistryString(string PobjValue)
        {
            try
            {
                if (string.IsNullOrEmpty(PobjValue))
                {
                    return;
                }

                // trim trailing "|" if it exists
                if (PobjValue.EndsWith("|"))
                {
                    PobjValue = PobjValue.Substring(0, PobjValue.Length - 1);
                }

                List<string> LobjItems = PobjValue.Split('|').ToList<string>();
                foreach (string LobjItem in LobjItems)
                {
                    //LstrRetVal += [0] LobjItem.DisplayName + ";" +
                    //              [1] LobjItem.HighlightColor.ToRGBColorString() + ";" +
                    //              [2] LobjItem.RecipientName + ";" +
                    //              [3] LobjItem.EntryId + ";" +
                    //              [4] LobjItem.ShowName.ToString() + ";" +
                    //              [5] LobjItem.Symbol + "|";
                    string[] LobjValues = LobjItem.Split(';');
                    this.Add(new ExtendedRecipient(LobjValues[2],
                                                   LobjValues[3],
                                                   bool.Parse(LobjValues[4]),
                                                   LobjValues[0],
                                                   LobjValues[1].FromRGBColorString(),
                                                   LobjValues[5]));
                }
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Failed to load recipient list form registry. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Imports a serialized list from file
        /// </summary>
        /// <param name="PstrFileName"></param>
        /// <returns></returns>
        public bool Import(string PstrFileName)
        {
            try
            {
                StreamReader LobjReader = new StreamReader(PstrFileName);
                XmlSerializer LobjSer = new XmlSerializer(this.GetType());
                this.Clear();
                this.AddRange((ExtendedRecipientList)LobjSer.Deserialize(LobjReader));
                LobjReader.Close();
                return true;
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to import the user list.");
                return false;
            }
        }

        /// <summary>
        /// Exports a serialized list to file
        /// </summary>
        /// <param name="PstrFilename"></param>
        /// <returns></returns>
        public bool Export(string PstrFilename)
        {
            try
            {
                StreamWriter LobjSw = new StreamWriter(PstrFilename);
                XmlSerializer LobjSer = new XmlSerializer(this.GetType());
                LobjSer.Serialize(LobjSw, this);
                LobjSw.Close();
                return true;
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to export the users.");
                return false;
            }
        }
    }
}
