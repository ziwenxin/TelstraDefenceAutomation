using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    public static class FileHelper
    {
        /// <summary>
        /// move a file to archive, this will delete the original file
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="filename"></param>
        /// <param name="extension"></param>
        public static void MoveFileToArchive(string savePath, string filename, bool deleteSrcFile)
        {
            string extension = ".xlsx";
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
            if (File.Exists(dstPath) && deleteSrcFile)
                File.Delete(dstPath);
            //copy the file to archive folder
            File.Copy(OriginalPath, dstPath);
            //delete the original file
            File.Delete(OriginalPath);
        }

        /// <summary>
        /// remove 'V2' from a str
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
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
