﻿using System;
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
        public static void SendEmail(string subject, string content)
        {
            LogHelper.AddToLog("Sending email...");
            //set address, subject and body of email
            string emailAddr = ConfigHelper._configDic["NotificationEmail"];

            string autoPath = ConfigHelper._configDic["AutomationPath"];
            string body = content;

            //set up a mail
            Application app = new Application();
            MailItem mail = (MailItem)app.CreateItem(OlItemType.olMailItem);

            mail.To = emailAddr;
            mail.Body = body;
            mail.Subject = subject;
            //set up the first account
            Accounts accs = app.Session.Accounts;
            Account acc = (Account)accs.First();
            mail.SendUsingAccount = acc;
            //send email
            mail.Send();
            LogHelper.AddToLog("Sending email completed");
        }

        /// <summary>
        /// this function will download all the attachments from a certain email of a certain account
        /// </summary>
        public static void DownloadAttachments()
        {
            //set up an application
            Application app = new Application();

            //set up account

            Accounts accs = app.Session.Accounts;
            string folderName = ConfigHelper._configDic["AttachmentEmail"];
            //get the inbox folder
            MAPIFolder inbox = app.Session.Folders[folderName].Folders["Inbox"];
            //get all the unread mails in today
            string restriction = "[Unread]=true";
            var items = inbox.Items.Restrict(restriction);

            //read settings
            string savePath = ConfigHelper._configDic["LocalSavePath"] + "\\SalesOrderHistory\\";
            if (!Directory.Exists(savePath))
                Directory.CreateDirectory(savePath);
            int totalSuppliers = Convert.ToInt32(ConfigHelper._configDic["TotalSuppliers"]);
            //read all the unread mails
            foreach (object o in items)
            {
                //it is a mail
                MailItem mi = o as MailItem;

                if (mi == null)
                    continue;
                for (int i = 0; i < totalSuppliers; i++)
                {
                    //if the email is send by the suppliers
                    string companyName = ConfigHelper._configDic["SupplierNames" + (i + 1)].Replace(" ", "").ToLower();
                    //replace double e with single e
                    if (companyName.Contains("ee"))
                        companyName = companyName.Replace("ee", "e");
                    if (mi.SenderEmailAddress.ToLower().Contains(companyName))
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
                                string rename = ConfigHelper._configDic["SupplierNames" + (i + 1)].ToUpper().Replace(" ", "_") + "_";
                                rename += extension;
                                attchment.SaveAsFile(savePath + rename);
                                //save the .xlsx directly
                                if (extension == ".xlsx")
                                    continue;
                                //save the .xls file as .xlsx

                                OfficeExcelHelper.SaveAs(savePath, rename);
                                //delete the original file
                                if (File.Exists(savePath + rename))
                                    File.Delete(savePath + rename);
                            }

                        }
                    }

                }
                //set it read
                mi.UnRead = false;

            }
        }
    }
}
