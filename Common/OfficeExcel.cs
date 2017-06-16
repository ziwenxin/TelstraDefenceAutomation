using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Common
{
    public static class OfficeExcel
    {
        public static void SaveAs(string path, string filename)
        {
            Application app=null;
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
                path=path.Replace("/", "\\");
                //get work book
                wb = appWorkbooks.Open(path + filename + ".xls");
                //save as
                wb.SaveAs(path + filename + ".xlsx", XlFileFormat.xlWorkbookDefault);
            }
            finally
            {
                wb.Close(0);
                app.Quit();
                //release
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(appWorkbooks);
                Marshal.ReleaseComObject(app);
            }


        }
    }
}
