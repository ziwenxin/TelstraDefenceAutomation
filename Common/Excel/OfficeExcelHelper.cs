using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;

namespace Common
{
    public static class OfficeExcelHelper
    {
        /// <summary>
        /// save a file from corrupted file to xlsx
        /// </summary>
        /// <param name="path"></param>
        /// <param name="filename"></param>
        public static void SaveAs(string path, string filename)
        {
            Application app = null;
            Workbook wb = null;
            Workbooks appWorkbooks = null;
            //get excel application
            try
            {
                app = new Application();
                if (app == null)
                {
                    throw new Exception("No Office Excel Installed");
                }
                //get work books
                appWorkbooks = app.Workbooks;
                path = path.Replace("/", "\\");
                //get work book
                if (!filename.EndsWith(".xls"))
                    wb = appWorkbooks.Open(path + filename + ".xls");
                else
                {
                    wb = appWorkbooks.Open(path + filename);
                    //remove .xls
                    filename = filename.Substring(0, filename.IndexOf(".xls"));
                }
                //save as
                wb.SaveAs(path + filename + ".xlsx", XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            finally
            {
                //release

                if (wb != null)
                {
                    wb.Close(0);
                    Marshal.ReleaseComObject(wb);

                }
                if (appWorkbooks != null)
                    Marshal.ReleaseComObject(appWorkbooks);
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

            }


        }
        /// <summary>
        /// change a sheet name and save it
        /// </summary>
        /// <param name="savepath"></param>
        /// <param name="filename"></param>
        /// <param name="srcName"></param>
        /// <param name="dstName"></param>
        public static void ChangeSheetName(string savepath, string filename, string srcName, string dstName)
        {
            //declare
            Application app = null;
            Workbook wb = null;
            Workbooks appWorkbooks = null;
            Sheets sheets = null;
            Worksheet sheet = null;
            //get excel application
            try
            {
                app = new Application();
                if (app == null)
                {
                    throw new Exception("No Office Excel Installed");
                }
                //disable alert
                app.DisplayAlerts = false;
                //get work books
                appWorkbooks = app.Workbooks;
                savepath = savepath.Replace("/", "\\");
                //get work book
                wb = appWorkbooks.Open(savepath + filename + ".xlsx");
                sheets = wb.Sheets;
                sheet = sheets[srcName];
                sheet.Name = dstName;

                //save as
                sheet.SaveAs(savepath + filename + ".xlsx", XlFileFormat.xlWorkbookDefault);
            }
            finally
            {
                //release
                if (sheet != null)
                    Marshal.ReleaseComObject(sheet);
                if (sheets != null)
                    Marshal.ReleaseComObject(sheets);
                if (wb != null)
                {

                    wb.Close(0);
                    Marshal.ReleaseComObject(wb);
                }
                if (appWorkbooks != null)
                    Marshal.ReleaseComObject(appWorkbooks);

                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }

            }
        }

        /// <summary>
        /// Delete a sheet from an excel
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="sheetName"></param>
        public static void DeleteASheet(string savePath, string sheetName)
        {
            //declare
            Application app = null;
            Workbook wb = null;
            Workbooks appWorkbooks = null;
            Sheets sheets = null;
            Worksheet sheet = null;
            //get excel application
            try
            {
                app = new Application();
                if (app == null)
                {
                    throw new Exception("No Office Excel Installed");
                }
                //disable alert
                app.DisplayAlerts = false;
                //get work books
                appWorkbooks = app.Workbooks;
                savePath = savePath.Replace("/", "\\");
                //get work book
                wb = appWorkbooks.Open(savePath);
                sheets = wb.Sheets;
                //delete the sheet
                if (sheets.Count > 1)
                    sheet = sheets[sheetName];
                if (sheet != null)
                {
                    sheet.Delete();
                    //save as
                    wb.SaveAs(savePath, XlFileFormat.xlWorkbookDefault);
                }


            }
            finally
            {
                //release
                if (sheet != null)
                    Marshal.ReleaseComObject(sheet);
                if (sheets != null)
                    Marshal.ReleaseComObject(sheets);
                if (wb != null)
                {

                    wb.Close(0);
                    Marshal.ReleaseComObject(wb);
                }
                if (appWorkbooks != null)
                    Marshal.ReleaseComObject(appWorkbooks);

                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
            }
        }

        public static void DeleteTopRows(int numRows, string fullpath)
        {
            //declare
            Application app = null;
            Workbook wb = null;
            Workbooks appWorkbooks = null;
            Sheets sheets = null;
            Worksheet sheet = null;
            Range rows = null;
            //get excel application
            try
            {
                app = new Application();
                if (app == null)
                {
                    throw new Exception("No Office Excel Installed");
                }
                //disable alert
                app.DisplayAlerts = false;
                //get work books
                appWorkbooks = app.Workbooks;
                fullpath = fullpath.Replace("/", "\\");
                //get work book
                wb = appWorkbooks.Open(fullpath);
                sheets = wb.Sheets;
                //get the first sheet
                if (sheets.Count > 0)
                    sheet = sheets[1];
                if (sheet != null)
                {
                    //delete first couple of rows

                    rows = sheet.Range["A1", "A" + numRows];
                    rows.EntireRow.Delete(XlDirection.xlUp);
                }

                //save as
                sheet.SaveAs(fullpath, XlFileFormat.xlWorkbookDefault);
            }

            finally
            {
                //release
                if (rows != null)
                    Marshal.ReleaseComObject(rows);

                if (sheet != null)
                    Marshal.ReleaseComObject(sheet);
                if (sheets != null)
                    Marshal.ReleaseComObject(sheets);
                if (wb != null)
                {

                    wb.Close(0);
                    Marshal.ReleaseComObject(wb);
                }
                if (appWorkbooks != null)
                    Marshal.ReleaseComObject(appWorkbooks);

                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
            }
        }
    }
}
