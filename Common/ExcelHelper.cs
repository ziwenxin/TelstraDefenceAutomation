using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
                    ISheet sheet = hssfWb.GetSheetAt(0);
                    return sheet;
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
        public static void SaveTo(ISheet sheet, string path)
        {
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
               sheet.Workbook.Write(fs);
            }
        }
    }
}


