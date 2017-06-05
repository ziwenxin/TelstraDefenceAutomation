using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
                XSSFWorkbook hssfWb = new XSSFWorkbook(fs);
                return hssfWb.GetSheetAt(0);
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
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write))
            {
                //create a new work book with the same sheet name
                XSSFWorkbook saveWorkbook = new XSSFWorkbook();
                saveWorkbook.CreateSheet(sheet.SheetName);
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
    }
}


