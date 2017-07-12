using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;

namespace Common
{
    public static class OutlookHelper
    {
        /// <summary>
        /// send a email to the address in config file, using the current outlook account
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="content"></param>
        /// <param name="configDic"></param>
        public static void SendEmail(string subject, string content, Dictionary<string, string> configDic)
        {
            //set address, subject and body of email
            string emailAddr = configDic["NotificationEmail"];

            string autoPath = configDic["AutomationPath"];
            string body = content;

            //set up a mail
            Application app = new Application();
            MailItem mail = (MailItem)app.CreateItem(OlItemType.olMailItem);

            mail.To = emailAddr;
            mail.Body = body;
            mail.Subject = subject;
            //set up account
            Accounts accs = app.Session.Accounts;
            Account acc = (Account)accs.First();
            mail.SendUsingAccount = acc;
            //send email
            mail.Send();
        }

        /// <summary>
        /// this function will download all the attachments from a certain email of a certain account
        /// </summary>
        /// <param name="configDic">the configuration data</param>
        public static void DownloadAttachments(Dictionary<string, string> configDic)
        {
            //set up an application
            Application app = new Application();

            //set up account
            Accounts accs = app.Session.Accounts;
            Account account = (Account)accs[configDic["AttachmentEmail"]];
            //get the inbox folder
            MAPIFolder Inbox = account.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            //get all the unread mails in today
            string restriction = "[ReceivedTime] >= '" + DateTime.Today.ToString("ddddd h:nn") +
                                 "'";
            var items = Inbox.Items.Restrict(restriction);

            //read settings
            string savePath = configDic["LocalSavePath"] + "\\"+ "SalesOrderHistory\\";
            if (!Directory.Exists(savePath))
                Directory.CreateDirectory(savePath);
            int totalSuppliers = Convert.ToInt32(configDic["TotalSuppliers"]);
            //read all the unread mails
            foreach (object o in items)
            {
                //it is a mail
                MailItem mi = o as MailItem;

                if (mi == null)
                    continue;
                //if it follows the rule
                for (int i = 0; i < totalSuppliers; i++)
                {
                    //if the email is send by the suppliers
                    if (mi.SenderEmailAddress.ToLower().Contains(configDic["SupplierNames" + (i + 1)].ToLower()))
                    {
                        //download all the attachments
                        foreach (var attchment in mi.Attachments)
                        {
                            if (attchment.FileName.Contains(".xls"))
                            {
                                //set extension
                                string extension = ".xls";
                                if (attchment.FileName.Contains(".xlsx"))
                                    extension = ".xlsx";
                                    //set rename
                                    string rename = configDic["SupplierName"].Replace(" ", "_") + "_" +DateTime.Today.ToString("dd-MM-yyyy");
                                rename += extension;
                                attchment.SaveAsFile(savePath + rename);
                                //save the .xlsx directly
                                if(extension==".xlsx")
                                    continue;
                                //save the .xls file as .xlsx

                                OfficeExcelHelper.SaveAs(savePath, rename);
                                //delete the original file
                                if(File.Exists(savePath+ rename))
                                    File.Delete(savePath+ rename);
                            }

                        }
                    }
                }

            }
        }
    }
}
