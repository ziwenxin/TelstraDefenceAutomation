using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Common;
using NPOI.SS.UserModel;

namespace TelstraDefenceAutomation
{
    public static class ExcelProcesser
    {
        /// <summary>
        /// process Meridian documents
        /// </summary>
        public static void ProcessMeridianExcels()
        {
            //get total meridian report numbers
            int totalWaitMilliSecs = 0;
            int totalReportNum = int.Parse(ConfigHelper._configDic["TotalMeridianDocuments"]);
            for (int i = 0; i < totalReportNum; i++)
            {
                //read from report
                string savePath = ConfigHelper._configDic["LocalSavePath"];
                string filename = ConfigHelper._configDic["OriginalFileName" + (i + 1)];
                savePath += "\\";
                string filepath = savePath + filename;
                string rename = ConfigHelper._configDic["Rename" + (i + 1)];
                string extension = ".xls";
                //check if the file exists
                if (!File.Exists(filepath + extension))
                {
                    throw new Exception(filepath + " is not downloaded");
                }
                //delete the top lines 
                ExcelHelper.ProcessInvalidExcel(savePath,filename,rename);

                //save the corrupted file as xlsx
                OfficeExcelHelper.SaveAs(savePath, rename);
                FileHelper.MoveFileToArchive(savePath, rename, false);
                //delete he original file
                if (File.Exists(savePath + filename + ".xls"))
                    File.Delete(savePath + filename + ".xls");

            }

        }


        /// <summary>
        /// process Toll documents
        /// </summary>
        public static void ProcessTollExcels()
        {
            //get total toll report numbers
            int totalWaitMilliSecs = 0;
            int totalReportNum = int.Parse(ConfigHelper._configDic["TotalTollDocuments"]);
            for (int i = 0; i < totalReportNum; i++)
            {
                //read from report
                string savePath = ConfigHelper._configDic["LocalSavePath"];
                string filename = ConfigHelper._configDic["TollDocumentName" + (i + 1)];
                savePath += "\\";
                string filepath = savePath + filename;
                //check if the file exists
                string extension = ".xlsx";

                if (!File.Exists(filepath + extension))
                {
                    throw new Exception(filepath + " is not downloaded");
                }
                int linesToBeDeleted = int.Parse(ConfigHelper._configDic["TollLinesToBeDeleted" + (i + 1)]);


                //do the archive
                FileHelper.MoveFileToArchive(savePath, filename, false);
                //save
                OfficeExcelHelper.DeleteTopRows(linesToBeDeleted, savePath + filename + extension);
                LogHelper.AddToLog(filename + " process completed");

            }


        }

        /// <summary>
        /// It will delete the top row of the excel
        /// </summary>
        public static void ProcessSucureExcel()
        {
            //get fullpath
            string savepath = ConfigHelper._configDic["LocalSavePath"] + "\\SalesOrderHistory\\";
            string companyName = ConfigHelper._configDic["SupplierNames4"].Replace(" ", "_");
            string fullPath = savepath + companyName + "_" + DateTime.Today.ToString("dd-MM-yyyy") + ".xlsx";
            if (File.Exists(fullPath))
            {
                //delete sheet 2
                OfficeExcelHelper.DeleteASheet(fullPath, "Sheet1");

                int lines = Convert.ToInt32(ConfigHelper._configDic["SupplierLinesToBeDeleted4"]);

                //delete the first line
                OfficeExcelHelper.DeleteTopRows(lines, fullPath);
            }
        }

        /// <summary>
        /// It will delete the top 4 rows of the excel
        /// </summary>
        public static void ProcessAvnetExcel()
        {
            //get fullpath
            string savepath = ConfigHelper._configDic["LocalSavePath"] + "\\SalesOrderHistory\\";
            string companyName = ConfigHelper._configDic["SupplierNames1"].Replace(" ", "_");
            string fullPath = savepath + companyName + "_" + DateTime.Today.ToString("dd-MM-yyyy") + ".xlsx";
            if (File.Exists(fullPath))
            {

                int lines = Convert.ToInt32(ConfigHelper._configDic["SupplierLinesToBeDeleted1"]);

                //delete the first 4 rows
                OfficeExcelHelper.DeleteTopRows(lines, fullPath);
            }
        }
    }
}
