using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoumangToolKit
{
   public static class fileListsAsync
    {


        public static async Task WalkTreeAsync(this List<FileInfo> fiList,string folderpath, bool recur = true, Action<FileInfo> doit=null)
        {
            // traverse the folder
            Func<string,Task> WalkDirectoryTree = null;
          WalkDirectoryTree = async delegate (string path)
              {
                  //fileInfo Allfile = new fileInfo();
                  DirectoryInfo dir = new DirectoryInfo(path);
                  FileInfo[] files = null;


                  try
                  {

                      files = dir.GetFiles();


                  }
                  //catch (UnauthorizedAccessException e)
                  catch (Exception e)
                  {
                      throw e;
                  }

                  if (files != null)
                  {

                      fiList.AddRange(files);

                      if(doit!=null)

                      {
                          foreach(var pp in files)
                          {
                              doit(pp);
                          }
                      }

                  }
                  // Now find all the subdirectories under this directory.
                  if (recur)
                  {
                      var subDirs = dir.GetDirectories();
                      // display = display + "\r\n";

                      foreach (System.IO.DirectoryInfo dirInfo in subDirs)
                      {
                          await WalkDirectoryTree(dirInfo.FullName);



                      }
                  }



              };
          await  WalkDirectoryTree(folderpath);
        }

        public static async Task WalkTreeAsync(string path,Action<FileInfo> doit)
        {
            List<FileInfo> fiList = new List<FileInfo>();

               await  fiList.WalkTreeAsync(path,doit:doit);
            

          
        }



    }
}
