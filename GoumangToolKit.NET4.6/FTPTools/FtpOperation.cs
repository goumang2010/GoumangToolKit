using System;
using System.Data;
using System.Configuration;
using System.Web;

using System.Net.Sockets;
using System.Text;
using System.Net;
using System.IO;
using System.Collections.Generic;

using System.Threading.Tasks;
using System.Linq;

namespace GoumangToolKit
{
 
    public  class FtpOperationAsync
    {
      string ftpServerIP;

        string ftpUserID;

        string ftpPassword;

        FtpWebRequest reqFTP;

        private void Connect(String path)//连接ftp
        {

            // 根据uri创建FtpWebRequest对象

            reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri(path));

            // 指定数据传输类型

            reqFTP.UseBinary = true;

            // ftp用户名和密码

            reqFTP.Credentials = new NetworkCredential(ftpUserID, ftpPassword);

        }

        public  FtpOperationAsync(string ftpServerIP, string ftpUserID, string ftpPassword)
        {
            this.ftpServerIP = ftpServerIP;

            this.ftpUserID = ftpUserID;

            this.ftpPassword = ftpPassword;
        }

        //都调用这个

        private async System.Threading.Tasks.Task<string[]> GetFileListAsync(string path, string WRMethods)//上面的代码示例了如何从ftp服务器上获得文件列表
        {
            string[] downloadFiles;
            StringBuilder result = new StringBuilder();
            try
            {
                Connect(path);

                reqFTP.Method = WRMethods;

                WebResponse response = await reqFTP.GetResponseAsync();

                StreamReader reader = new StreamReader(response.GetResponseStream(), System.Text.Encoding.Default);//中文文件名

                string line = reader.ReadLine();

                while (line != null)
                {

                    result.Append(line);

                    result.Append("\n");

                    line = reader.ReadLine();

                }

                // to remove the trailing '\n'

                result.Remove(result.ToString().LastIndexOf('\n'), 1);

                reader.Close();

                response.Close();

                return result.ToString().Split('\n');

            }

            catch (Exception ex)
            {
              

                downloadFiles = null;

                return downloadFiles;
            }
        }

        public async System.Threading.Tasks.Task<string[]> GetFileListAsync(string path)//上面的代码示例了如何从ftp服务器上获得文件列表
        {
            return await GetFileListAsync("ftp://" + ftpServerIP + "/" + path, WebRequestMethods.Ftp.ListDirectory);
        }

        public async System.Threading.Tasks.Task<string[]> GetFileListAsync()//上面的代码示例了如何从ftp服务器上获得文件列表
        {
           return await GetFileListAsync("ftp://" + ftpServerIP + "/", WebRequestMethods.Ftp.ListDirectory);
        }

        public async void UploadAsync(string filename, String directory, String newFileName) //上面的代码实现了从ftp服务器上载文件的功能
        {

            FileInfo fileInf = new FileInfo(filename);

            string uri = "ftp://" + ftpServerIP + "/" + directory + "/" + newFileName;

            if (!await DirectoryIsExistAsync(directory))
            {
                await MakeDirAsync(directory);
            }

            Connect(uri);//连接        

            // 默认为true，连接不会被关闭

            // 在一个命令之后被执行

            reqFTP.KeepAlive = false;

            // 指定执行什么命令

            reqFTP.Method = WebRequestMethods.Ftp.UploadFile;

            // 上传文件时通知服务器文件的大小

            reqFTP.ContentLength = fileInf.Length;
            // 缓冲大小设置为kb 
            int buffLength = 2048;
            byte[] buff = new byte[buffLength];

            int contentLen;

            // 打开一个文件流(System.IO.FileStream) 去读上传的文件

            FileStream fs = fileInf.OpenRead();

            try
            {

                // 把上传的文件写入流

                Stream strm =await reqFTP.GetRequestStreamAsync();

                // 每次读文件流的kb

                contentLen = fs.Read(buff, 0, buffLength);

                // 流内容没有结束

                while (contentLen != 0)
                {
                    // 把内容从file stream 写入upload stream 
                    strm.Write(buff, 0, contentLen);
                    contentLen = fs.Read(buff, 0, buffLength);

                }

                // 关闭两个流

                strm.Close();

                fs.Close();

            }

            catch (Exception ex)
            {

                throw ex;
            }

        }

        public bool Download(string filePath, string fileName, out string errorinfo)
        {
            try
            {
                String onlyFileName = Path.GetFileName(fileName);

                string newFileName = filePath + "\\" + onlyFileName;

                if (File.Exists(newFileName))
                {

                    errorinfo = string.Format("本地文件{0}已存在,无法下载", newFileName);
                    return false;
                }
                string url = "ftp://" + ftpServerIP + "/" + fileName;
                Connect(url);//连接 
                reqFTP.Credentials = new NetworkCredential(ftpUserID, ftpPassword);
                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                Stream ftpStream = response.GetResponseStream();
                long cl = response.ContentLength;
                int bufferSize = 2048;
                int readCount;
                byte[] buffer = new byte[bufferSize];
                readCount = ftpStream.Read(buffer, 0, bufferSize);

                FileStream outputStream = new FileStream(newFileName, FileMode.Create);
                while (readCount > 0)
                {
                    outputStream.Write(buffer, 0, readCount);
                    readCount = ftpStream.Read(buffer, 0, bufferSize);
                }
                ftpStream.Close();
                outputStream.Close();
                response.Close();

                errorinfo = "";

                return true;

            }

            catch (Exception ex)
            {
                errorinfo = string.Format("因{0},无法下载", ex.Message);

                return false;

            }

        }

        //删除文件

        public void DeleteFileName(string fileName)
        {
            try
            {
                FileInfo fileInf = new FileInfo(fileName);

                string uri = "ftp://" + ftpServerIP + "/" + fileInf.Name;

                Connect(uri);//连接        

                // 默认为true，连接不会被关闭

                // 在一个命令之后被执行

                reqFTP.KeepAlive = false;

                // 指定执行什么命令

                reqFTP.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                response.Close();

            }

            catch (Exception ex)
            {

                throw ex;

            }

        }

        //创建目录

        public async Task MakeDirAsync(string dirName)
        {
            try
            {
                string uri = "ftp://" + ftpServerIP + "/" + dirName;

                Connect(uri);//连接     

                reqFTP.Method = WebRequestMethods.Ftp.MakeDirectory;

               var response =await reqFTP.GetResponseAsync();

                response.Close();

            }

            catch (Exception ex)
            {
                throw ex;
             
            }

        }

        //删除目录

        public void delDir(string dirName)
        {
            try
            {
                string uri = "ftp://" + ftpServerIP + "/" + dirName;

                Connect(uri);//连接     

                reqFTP.Method = WebRequestMethods.Ftp.RemoveDirectory;

                FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();

                response.Close();

            }

            catch (Exception ex)
            {

                throw ex;

            }

        }

        //获得文件大小

        public async Task<long> GetFileSizeAsync(string filename)
        {


            long fileSize = 0;

            try
            {

                FileInfo fileInf = new FileInfo(filename);

                string uri = "ftp://" + ftpServerIP + "/" + fileInf.Name;

                Connect(uri);//连接     

                reqFTP.Method = WebRequestMethods.Ftp.GetFileSize;

                FtpWebResponse response = (FtpWebResponse)await reqFTP.GetResponseAsync();

                fileSize = response.ContentLength;

                response.Close();

            }

            catch (Exception ex)
            {

                throw ex;

            }

            return fileSize;

        }

        //文件改名

        public async void RenameAsync(string currentFilename, string newFilename)
        {
            try
            {
                FileInfo fileInf = new FileInfo(currentFilename);

                string uri = "ftp://" + ftpServerIP + "/" + fileInf.Name;

                Connect(uri);//连接

                reqFTP.Method = WebRequestMethods.Ftp.Rename;

                reqFTP.RenameTo = newFilename;


                FtpWebResponse response = (FtpWebResponse)await reqFTP.GetResponseAsync();

                //Stream ftpStream = response.GetResponseStream();

                //ftpStream.Close();

                response.Close();

            }

            catch (Exception ex)
            {

                throw ex;

            }

        }

        //获得文件明晰

        public async Task<string[]> GetFilesDetailListAsync()
        {

            return await GetFileListAsync("ftp://" + ftpServerIP + "/", WebRequestMethods.Ftp.ListDirectoryDetails);

        }

        //Recursion for operate files
        public async Task<bool> WalktreeFTP(string Rootpath,Action<string,string> fileOp)
        {
            Func<string,Task> recursion = null;
            recursion = async delegate (string path)
              {
                  string[] ifdictinary=new string[1];
                  string[] allfile=new string[1];

                      ifdictinary = await GetFilesDetailListAsync(path);


                      allfile = await GetFileListAsync(path);



                  for (int i = 4; i < ifdictinary.Count(); i++)
                  {
                      int m = i - 2;
                      string filename = allfile[m];
                      string kk = ifdictinary[i];
                      if (kk.Substring(0, 1) == "d")
                      {
                          //如果为目录，则递归遍历
                        await  recursion(path + "/" + filename);
                      }
                      //If it is a file
                      else if (kk.Substring(0, 1) == "-")
                      {
                          string folderpath = "ftp://" + ftpServerIP + "/" + path + "/";
                          fileOp(filename, folderpath);


                      }

                  }
              };
           await  recursion(Rootpath);
           return true;

        }
        public async Task<string[]> GetFilesDetailListAsync(string path)
        {

            return await GetFileListAsync("ftp://" + ftpServerIP + "/" + path, WebRequestMethods.Ftp.ListDirectoryDetails);
        }

        /// <summary>
        /// 检测目录是否存在
        /// </summary>
        /// <param name="dirName"></param>
        /// <returns>false不存在，true存在</returns>
        public async Task<bool> DirectoryIsExistAsync(string dirName)
        {
            string[] value = await GetFileListAsync(dirName);
            if (value == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }
}
