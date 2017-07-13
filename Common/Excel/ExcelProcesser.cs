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
                    if (!File.Exists(filepath + ".xls"))
                        throw new Exception(filepath + " is not downloaded");
                }
                //process the file by string
                ExcelHelper.ProcessInvalidExcel(savePath, filename, rename);
                //save the corrupted file as xlsx
                OfficeExcelHelper.SaveAs(savePath, rename);
                FileHelper.MoveFileToArchive(savePath, rename, false);
                //delete he original file
                if (File.Exists(savePath + rename + ".xls"))
                    File.Delete(savePath + rename + ".xls");

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
                    if (!File.Exists(filepath + ".xls"))
                        throw new Exception(filepath + " is not downloaded");
                }
                int linesToBeDeleted = int.Parse(ConfigHelper._configDic["LinesToBeDeleted" + (i + 1)]);

                //use library to read an excel file
                ISheet reportsheet = ExcelHelper.ReadExcel(filepath + extension);

                //do the archive
                FileHelper.MoveFileToArchive(savePath, filename, true);
                //save
                ExcelHelper.SaveTo(reportsheet, filepath + ".xlsx", linesToBeDeleted);
                LogHelper.AddToLog(filename + " process completed");

            }


        }

        public static void ProcessSucureExcel()
        {
            //get fullpath
            string savepath = ConfigHelper._configDic["LocalSavePath"] + "\\SalesOrderHistory\\";
            string companyName = ConfigHelper._configDic["SupplierNames4"].Replace(" ","_");
            string fullPath = savepath + companyName+"_" + DateTime.Today.ToString("dd-MM-yyyy")+".xlsx";
            if (File.Exists(fullPath))
            {
                //delete sheet 2
                OfficeExcelHelper.DeleteASheet(fullPath, "Sheet1");

                //delete the first line
                ISheet sheet = ExcelHelper.ReadExcel(fullPath);
                ExcelHelper.SaveTo(sheet, fullPath, 1);
            }
        }
    }
}
