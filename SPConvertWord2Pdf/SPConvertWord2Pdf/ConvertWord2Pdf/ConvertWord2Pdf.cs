using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using Microsoft.Office.Word.Server.Conversions;

namespace SPConvertWord2Pdf.ConvertWord2Pdf
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class ConvertWord2Pdf : SPItemEventReceiver
    {
        /// <summary>
        /// An item was updated.
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            base.ItemUpdated(properties);
            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    var spItem = properties.ListItem;
                    if (bool.Parse(spItem["Converter"].ToString()))
                    {
                        if (properties.ListItem.Name.ToLower().Contains(".docx") || properties.ListItem.Name.ToLower().Contains(".doc"))
                        {
                            string wordFile = string.Empty;
                            string pdfFile = string.Empty;
                            string wordAutomationServiceName = "Word Automation Services";

                            wordFile = properties.WebUrl + "/" + properties.ListItem.Url;

                            if (properties.ListItem.Name.ToLower().Contains(".docx"))
                            {
                                pdfFile = string.Format("{0}/DocPdf/{1}", properties.WebUrl, properties.ListItem.Name.ToLower().Replace(".docx", ".pdf"));
                            }
                            else
                            {
                                pdfFile = string.Format("{0}/DocPdf/{1}", properties.WebUrl, properties.ListItem.Name.ToLower().Replace(".doc", ".pdf"));
                            }

                            SyncConverter sc = new SyncConverter(wordAutomationServiceName);
                            sc.UserToken = properties.Web.CurrentUser.UserToken;
                            sc.Settings.UpdateFields = true;
                            sc.Settings.OutputFormat = SaveFormat.PDF;

                            ConversionItemInfo info = sc.Convert(wordFile, pdfFile);
                            while (info.NotStarted)
                            {
                                spItem["StatusConversao"] = "NotStarted";
                                spItem.Update();
                            }
                            while (info.InProgress)
                            {
                                spItem["StatusConversao"] = "InProgress";
                                spItem.Update();
                            }

                            if (info.Succeeded)
                            {
                                spItem["StatusConversao"] = "Succeeded";
                                spItem["StartTime"] = info.StartTime;
                                spItem["CompleteTime"] = info.CompleteTime;
                            }
                            else if (info.Canceled)
                            {
                                spItem["StatusConversao"] = "Canceled";
                                spItem["Informacao"] = string.Format("ErrorCode:{0}#;Message{1}", info.ErrorCode, info.ErrorMessage);
                            }
                            else if (info.Failed)
                            {
                                spItem["StatusConversao"] = "Failed";
                                spItem["Informacao"] = string.Format("ErrorCode:{0}#;Message{1}", info.ErrorCode, info.ErrorMessage);
                            }

                            spItem.Update();
                        }
                    }
                });
            }
            catch (SPException ex)
            {
                throw ex;
            }
        }
    }
}