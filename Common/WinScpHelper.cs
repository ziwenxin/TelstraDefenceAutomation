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
        public static void UploadFiles(Dictionary<string, string> configDic, ref StringBuilder sb)
        {

            //setup session options
            SessionOptions options = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = configDic["HostName"],
                UserName = configDic["WinScpUsername"],
                Password = configDic["WinScpPassword"],
                SshHostKeyFingerprint = configDic["FingerPrint"],
            };

            using (Session session = new Session())
            {
                //connect
                session.Open(options);

                //upload files
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;

                //get path
                string localPath = configDic["LocalSavePath"];
                string remotePath = configDic["RemoteSavePath"];
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
                    string msg = string.Format("Upload of {0} successed", transfer.FileName);
                    sb.Append(msg + "\r\n");
                    Console.WriteLine(msg);
                }
                

                //upload supplier documents
                UploadSupplierDocuments(session, configDic, ref sb);

            }

        }

        private static void UploadSupplierDocuments(Session session, Dictionary<string, string> configDic, ref StringBuilder sb)
        {
            //transfer options
            TransferOptions transferOptions = new TransferOptions();
            transferOptions.TransferMode = TransferMode.Binary;
            //results
            TransferOperationResult operationResult = null;
            RemovalOperationResult removalOperationResult = null;
            //get local path and total suppliers
            string localPath = configDic["LocalSavePath"] + "\\SalesOrderHistory\\";
            string remotePath = configDic["RemoteSavePath"] + "/SalesOrderHistory/";
            int totalSuppliers = Convert.ToInt32(configDic["TotalSuppliers"]);

            for (int i = 0; i < totalSuppliers; i++)
            {
                //check if the updated file exists
                string supplierFileName = configDic["SupplierNames" + (i + 1)];
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

                    string msg = string.Format("Upload of {0} successed", operationResult.Transfers.First().FileName);
                    sb.Append(msg + "\r\n");
                    Console.WriteLine(msg);
                }

            }


        }
    }
}
