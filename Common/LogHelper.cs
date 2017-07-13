using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public static class LogHelper
    {
        private static StringBuilder sb=new StringBuilder();


        /// <summary>
        /// print message to output window and log the message
        /// </summary>
        /// <param name="msg"></param>
        public static void AddToLog(string msg)
        {
            sb.Append(msg + "\r\n");
            Console.WriteLine(msg);
        }

        public static void WriteLog(string filename)
        {
            File.WriteAllText(filename, sb.ToString());
        }
    }
}
