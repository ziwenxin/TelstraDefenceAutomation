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




  


    }
}


