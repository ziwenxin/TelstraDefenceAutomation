using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinSCP;

namespace Common
{
    public static class WinScpHelper
    {
        /// <summary>
        /// upload all the files to the server using WinScp
        /// </summary>
        public static void UploadFiles()
        {
            LogHelper.AddToLog("Start to upload files to server...");
            //setup session options
            SessionOptions options = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = ConfigHelper._configDic["HostName"],
                UserName = ConfigHelper._configDic["WinScpUsername"],
                Password = ConfigHelper._configDic["WinScpPassword"],
                SshHostKeyFingerprint = ConfigHelper._configDic["FingerPrint"],
            };

            using (Session session = new Session())
            {
                //connect
                session.Open(options);

                //upload files
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;

                //get path
                string localPath = ConfigHelper._configDic["LocalSavePath"];
                string remotePath = ConfigHelper._configDic["RemoteSavePath"];
                localPath += "\\";
                remotePath += "/";
                //change the '/' to '\'
                localPath = localPath.Replace("/", "\\");

                //remove all files before uploading
                session.RemoveFiles(remotePath + "*.xlsx");
                //upload the files into server,delete the files in the local
                TransferOperationResult operationResult =
                    session.PutFiles(localPath + "*.xlsx", remotePath, true, transferOptions);
                //throw any error
                operationResult.Check();

                //print result
                foreach (TransferEventArgs transfer in operationResult.Transfers)
                {
                    LogHelper.AddToLog(string.Format("Upload of {0} successes", transfer.FileName));
                }
                

                //upload supplier documents
                UploadSupplierDocuments(session);

                LogHelper.AddToLog("Upload all completed");
            }

        }

        private static void UploadSupplierDocuments(Session session)
        {
            //transfer options
            TransferOptions transferOptions = new TransferOptions();
            transferOptions.TransferMode = TransferMode.Binary;
            //results
            TransferOperationResult operationResult = null;
            RemovalOperationResult removalOperationResult = null;
            //get local path and total suppliers
            string localPath = ConfigHelper._configDic["LocalSavePath"] + "\\SalesOrderHistory\\";
            string remotePath = ConfigHelper._configDic["RemoteSavePath"] + "/SalesOrderHistory/";
            int totalSuppliers = Convert.ToInt32(ConfigHelper._configDic["TotalSuppliers"]);

            for (int i = 0; i < totalSuppliers; i++)
            {
                //check if the updated file exists
                string supplierFileName = ConfigHelper._configDic["SupplierNames" + (i + 1)];
                string pattern = supplierFileName + "*.xlsx";
                //if it needs to be updated
                foreach (var file in session.EnumerateRemoteFiles(remotePath, pattern, EnumerationOptions.None))
                {
                    //delete original 1
                    session.RemoveFiles(file.FullName);

                }
                //upload new one
                operationResult = session.PutFiles(localPath + pattern, remotePath, true, transferOptions);

                //throw any error
                operationResult.Check();

                //print result
                if (operationResult.Transfers.Count >0)
                {

                    LogHelper.AddToLog(string.Format("Upload of {0} successes", operationResult.Transfers.First().FileName));
                }

            }


        }
    }
}
