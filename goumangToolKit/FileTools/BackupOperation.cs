using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace GoumangToolKit
{
   public static class BackupOperation
    {


        public static void backupfile(string filepath)
        {
            string filename = filepath.Split('\\').Last();
            string folderpath = filepath.Replace(filename, "");
            string backupfolder = folderpath + @"backup\";
            localMethod.creatDir(backupfolder);

            if (File.Exists(filepath))
            {
                File.Copy(filepath, backupfolder + filename.Replace(".", "_backup_" + DateTime.Now.ToString("yyyy-MM-dd_hh-mm") + "."), true);
            }



        }
        public static void backupfolder(string newfoldername)
        {
            //Remove all other files 
            List<FileInfo> oldfiles = new List<FileInfo>();
            oldfiles.WalkTree(newfoldername, false);
            string backupfolder = newfoldername + "backup_" + DateTime.Now.ToString("yyyy-MM-dd_hh-mm");
            localMethod.creatDir(backupfolder);
            oldfiles.moveto(backupfolder);
        }


        public static int CleanBackup(string startPath)
        {
            var delarray = new List<FileInfo>();

            delarray.WalkTree(startPath);


            var todel = from pp in delarray
                        let fn = pp.Name
                        let fna = fn.Split('_').Reverse()
                        let d = (DateTime.Parse(fna.ElementAt(1)) - DateTime.Now).Days
                        where fna.Contains("backup") && (fna.Count() > 3) && Math.Abs(d) > 30
                        select pp;

            foreach (var dd in todel)
            {
                dd.Delete();

            }
            return todel.Count();

        }

    }
}
