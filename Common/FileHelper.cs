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
        /// add time stamps to all the files in a directory
        /// </summary>
        /// <param name="path"></param>
        public static void AddTimeStamps(string path)
        {
            DirectoryInfo di=new DirectoryInfo(path);
            foreach (var file in di.GetFiles())
            {
                //make each of the file to upper name
                if (file.Name.Contains(DateTime.Today.ToString("dd-MM-yyyy")))
                    continue;
                string upperName =file.DirectoryName+"\\"+ file.Name.Substring(0, file.Name.IndexOf("."))+ DateTime.Today.ToString("dd-MM-yyyy") + file.Extension;
                File.Move(file.FullName,upperName);
            }
        }
        /// <summary>
        /// this function will delete all the archives older than 3 months
        /// </summary>
        /// <param name="path">save path</param>
        public static void DeleteOldArchive(string path)
        {
            //get directory info
            DirectoryInfo di = new DirectoryInfo(path);
            foreach (var directory in di.GetDirectories())
            {
                DateTime dt = Convert.ToDateTime(directory.Name);
                if (DateTime.Now.Subtract(dt).Days >= 90)
                {
                    directory.Delete();
                }
            }
        }
        /// <summary>
        /// delete all the files in the local save path, excluding the folder
        /// </summary>
        /// <param name="path">the folder wants to delete all the files</param>
        public static void DeleteAllFiles(string path)
        {
            LogHelper.AddToLog("Deleting all previous files...");
            //delete reports
            DirectoryInfo di = new DirectoryInfo(path);
            foreach (FileInfo fileInfo in di.GetFiles())
            {
                fileInfo.Delete();
            }
            LogHelper.AddToLog("Delete completed");
        }
        /// <summary>
          /// move a file to archive, this will delete the original file
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="filename"></param>
        /// <param name="deleteSrcFile">whether to delete source file or not</param>
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
            if (File.Exists(dstPath))
                File.Delete(dstPath);
            //copy the file to archive folder
            File.Copy(OriginalPath, dstPath);
            //delete the original file
            if (deleteSrcFile)
                File.Delete(OriginalPath);
        }

        /// <summary>
        /// remove 'V2' from a str
        /// </summary>
        /// <param name="str"> the sources string</param>
        /// <returns>the string without v2</returns>
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

        /// <summary>
        /// find the newest version of file in the share folder
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="filename"></param>
        /// <returns>the newest file name</returns>
        public static string GetNewestFileName(string filepath, string filename)
        {

            //get directory info
            DirectoryInfo directory = new DirectoryInfo(filepath);
            //get the latest file
            return directory.GetFiles(filename + "*.xlsx").OrderByDescending(f => f.LastWriteTime).First().Name;
        }


        /// <summary>
        /// Delete a file if exists
        /// </summary>
        /// <param name="savepath"></param>
        /// <param name="filename"></param>
        public static void DeleteFile(string savepath, string filename)
        {
            string fullPath = savepath + filename + ".xls";
            if (File.Exists(fullPath))
                File.Delete(fullPath);
        }
    }
}
