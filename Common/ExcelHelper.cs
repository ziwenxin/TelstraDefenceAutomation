using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Common
{
    public static class ExcelHelper
    {

        public static ISheet ReadExcel(string filepath)
        {
            using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook xssfWb = new XSSFWorkbook(fs);
                return xssfWb.GetSheetAt(0);



            }
        }

        //delete a number of rows
        public static void DeleteRows(ISheet sheet, int num)
        {

            //remove rows
            for (int i = 0; i < num; i++)
            {
                sheet.RemoveRow(sheet.GetRow(i));
            }
            //move the remain up
            sheet.ShiftRows(num, sheet.LastRowNum, -1);

        }

        //save a excel
        public static void SaveTo(ISheet sheet, string path, int linesToBeDeleted)
        {
            path = RemoveV2(path);

            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            {

                //create a new work book with the same sheet name
                XSSFWorkbook saveWorkbook = new XSSFWorkbook();
                saveWorkbook.CreateSheet(RemoveV2(sheet.SheetName));
                ISheet newSheet = saveWorkbook.GetSheetAt(0);

                //copy data row by row
                for (int i = 0; i < sheet.LastRowNum - linesToBeDeleted; i++)
                {
                    IRow newRow = newSheet.CreateRow(i);
                    CopyRow(newRow, sheet.GetRow(i + linesToBeDeleted));

                }

                //autosize the columns
                for (int i = 0; i < sheet.GetRow(0).PhysicalNumberOfCells; i++)
                {
                    newSheet.AutoSizeColumn(i);

                }

                saveWorkbook.Write(fs);
            }

        }

        public static void Save(ISheet sheet,string filename)
        {
            using (FileStream fs=new FileStream(filename, FileMode.OpenOrCreate,FileAccess.ReadWrite))
            {
                sheet.Workbook.Write(fs);
            }
        }
        private static void CopyRow(IRow newRow, IRow srcRow)
        {
            for (int i = 0; i < srcRow.LastCellNum; i++)
            {
                ICell newCell = newRow.CreateCell(i);
                CopyCell(newCell, srcRow.GetCell(i));
            }
        }

        //copy data in a cell
        private static void CopyCell(ICell newCell, ICell srcCell)
        {

            //copy due to cell type
            CellType srcCellType = srcCell.CellType;
            if (srcCellType == CellType.Numeric)
            {
                //it is date value
                if (DateUtil.IsCellDateFormatted(srcCell))
                {
                    newCell.SetCellValue(srcCell.DateCellValue.ToString("d"));
                }
                else
                    newCell.SetCellValue(srcCell.NumericCellValue);
            }
            else if (srcCellType == CellType.String)
            {

                newCell.SetCellValue(srcCell.RichStringCellValue);
            }
            else if (srcCellType == CellType.Blank)
            {
                // nothing21
            }
            else if (srcCellType == CellType.Boolean)
            {
                newCell.SetCellValue(srcCell.BooleanCellValue);
            }
            else if (srcCellType == CellType.Error)
            {
                newCell.SetCellErrorValue(srcCell.ErrorCellValue);
            }
            else if (srcCellType == CellType.Formula)
            {
                newCell.SetCellFormula(srcCell.CellFormula);
            }
            else
            { // nothing29
            }
        }

        //process excel file with invalid header
        public static void ProcessInvalidExcel(string path, string filename,string rename)
        {
            //read all the lines
            int tableTagCount = 0;
            string[] allLines = File.ReadAllLines(path + filename + ".xls");
            //transfer the array to a list
            List<string> strList = new List<string>(allLines);
            for (int i = 0; i < strList.Count; i++)
            {
                string line = strList[i];
                //delete the image
                if (line.Contains("<table>") && tableTagCount == 0)
                {
                    //delete the belowing 5 lines
                    for (int j = 0; j < 2; j++)
                    {
                        strList.RemoveAt(i+j);

                    }
                    tableTagCount++;
                }
                //delete the line above content
                else if (line.Contains("<table>"))
                {
                    //find the last idx of "<table>"
                    int lastIdx = line.LastIndexOf("<table>");
                    //remove all the text before it
                    strList[i] = strList[i].Substring(lastIdx - "<table>".Length + "<table>".Length);
                    break;
                }
            }
            //make an archive
            MoveFileToArchive(path, filename, ".xls");
            //save the text into a new file
            File.WriteAllLines(path + rename + ".xls", strList.ToArray());

        }


        public static void MoveFileToArchive(string savePath, string filename, string extension)
        {
            //save set archivepath and archive file name
            string archivePath = savePath + "Archive/";
            if (!Directory.Exists(archivePath))
                Directory.CreateDirectory(archivePath);
            //set date format
            string dateStr = DateTime.Today.ToString("d");
            dateStr = dateStr.Replace("/", "-");
            //set a data folder in the archive folder
            archivePath += dateStr + "/";
            if (!Directory.Exists(archivePath))
                Directory.CreateDirectory(archivePath);
            string archiveFilename = filename + " " + dateStr;
            //set destination path and original path
            string OriginalPath = savePath + filename + extension;
            string dstPath = archivePath + archiveFilename + extension;
            //if the archive file exists, delete it
            if (File.Exists(dstPath))
                File.Delete(dstPath);
            //copy the file to archive folder
            File.Copy(OriginalPath, dstPath);
            //delete the original file
            File.Delete(OriginalPath);
        }

        public static string RemoveV2(string str)
        {
            //delete v2 if the filename contains
            string lowerPath = str.ToLower();
            if (lowerPath.Contains("v2"))
            {
                //find the index of v2
                int index = lowerPath.IndexOf("v2");
                //remove it
                str = str.Substring(0, index - 1) + str.Substring(index + 2, str.Length - index - 2);
            }
            return str;
        }
    }
}


