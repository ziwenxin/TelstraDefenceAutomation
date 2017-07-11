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
        /// <summary>
        /// read an excel
        /// </summary>
        /// <param name="filepath"></param>
        /// <returns>the data sheet</returns>
        public static ISheet ReadExcel(string filepath)
        {
            using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook xssfWb = new XSSFWorkbook(fs);
                return xssfWb.GetSheetAt(0);


            }
        }

        /// <summary>
        /// delete a number of rows from a sheet
        /// </summary>
        /// <param name="sheet">the sheet needed to be change</param>
        /// <param name="num">the number needed to delete</param>
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

        /// <summary>
        /// save a file to another path, delete several lines in it
        /// </summary>
        /// <param name="sheet">the save sheet</param>
        /// <param name="path"></param>
        /// <param name="linesToBeDeleted">the number of rows needs to be deleted</param>
        public static void SaveTo(ISheet sheet, string path, int linesToBeDeleted)
        {
            path = FileHelper.RemoveV2(path);

            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            {

                //create a new work book with the same sheet name
                XSSFWorkbook saveWorkbook = new XSSFWorkbook();
                saveWorkbook.CreateSheet(FileHelper.RemoveV2(sheet.SheetName));
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

        /// <summary>
        /// save a excel file
        /// </summary>
        /// <param name="sheet">the save sheet</param>
        /// <param name="filename"></param>
        public static void Save(ISheet sheet,string filename)
        {
            using (FileStream fs=new FileStream(filename, FileMode.OpenOrCreate,FileAccess.ReadWrite))
            {
                sheet.Workbook.Write(fs);
            }
        }
        /// <summary>
        /// copy a row
        /// </summary>
        /// <param name="newRow"></param>
        /// <param name="srcRow"></param>
        private static void CopyRow(IRow newRow, IRow srcRow)
        {
            for (int i = 0; i < srcRow.LastCellNum; i++)
            {
                ICell newCell = newRow.CreateCell(i);
                CopyCell(newCell, srcRow.GetCell(i));
            }
        }

        /// <summary>
        /// copy a cell
        /// </summary>
        /// <param name="newCell"></param>
        /// <param name="srcCell"></param>
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
            { // nothing
            }
        }

        /// <summary>
        /// process the corrupted file and rename it
        /// </summary>
        /// <param name="path"></param>
        /// <param name="filename">original name</param>
        /// <param name="rename"></param>
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
            //save the text into a new file
            File.WriteAllLines(path + rename + ".xls", strList.ToArray());

        }




  

        /// <summary>
        /// change a sheet name from srcName to dstName
        /// </summary>
        /// <param name="sheet">it could be any sheet inside of the workbook</param> 
        /// <param name="srcName"></param>
        /// <param name="dstName"></param>
        public static ISheet ChangeSheetName(ISheet sheet, string srcName, string dstName)
        {
            //get work book
            IWorkbook workbook = sheet.Workbook;
            //get sheet idx by original name
            int idx = workbook.GetSheetIndex(srcName);
            //set its name
            workbook.SetSheetName(idx,dstName);
            return workbook.GetSheetAt(idx);
        }

    }
}


