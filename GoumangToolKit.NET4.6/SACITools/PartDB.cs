using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Bson;
using MongoDB.Driver;

namespace GoumangToolKit
{
    public class PartDB
    {

        MongoMethod model;


       public   PartDB( string col)
        {
            string client = (string)(localMethod.GetConfigValue("MONGO_URI", "PartDBCfg.py"));
            string database = (string)(localMethod.GetConfigValue("MONGO_DATABASE", "PartDBCfg.py"));
            //string collection = (string)(localMethod.GetConfigValue("PART_MONGO_COLNAME", "PartDBCfg.py"));
            model = new MongoMethod(client, database, col);

        }


      public async Task<bool> UpdatePartDB(string path)
        {
            List<string> filenameQueue = new List<string>();
            Action<FileInfo> syncDB = delegate (FileInfo fi)
              {
                  if (!filenameQueue.Contains(fi.Name))
                  {
                  //Find the same name file
                  var filter = new BsonDocument("FileName", fi.Name);
                  var check =  model.collection.Find(filter);

                  if (check.Count() == 0)
                  {
                      string rev = "";
                      int mindex = fi.Name.LastIndexOf('.');
                     if( mindex>=2)
                      {
                          rev = fi.Name.Substring(mindex - 2, 2);
                      }
                      BsonDocument bd = new BsonDocument {
                    { "FileName",fi.Name },
                    { "Rev",rev },
                    { "Extention",fi.Extension},
                     { "FilePath",fi.FullName },
                      { "InsertDate",DateTime.Now.ToShortDateString()}
                  };
                      model.collection.InsertOne(bd);
                  }
                  else

                  {
                      //检查地址是否存在，若不存在，则更新记录
                      var item = check.First();
                      if (!File.Exists(item["FilePath"].AsString))
                      {
                          var updatestr = new BsonDocument {
                           { "$set",
                           new BsonDocument { { "FilePath", fi.FullName },
                           { "InsertDate", DateTime.Now.ToShortDateString() }}

                          } };
                          model.collection.UpdateOne(filter, updatestr);
                      }
                  }
                      filenameQueue.Add(fi.Name);
                  }
              };
            try
            {

              await fileListsAsync.WalkTreeAsync(path, syncDB);
                return true;
            }
            catch
            {
                return false;
            }


        }

       public async Task<bool> UpdateFTPPartDB(string path)
        {
            List<string> filenameQueue = new List<string>();
            string ftp_addr = (string)(localMethod.GetConfigValue("FTP_ADDR", "PartDBCfg.py"));
            string ftp_user= (string)(localMethod.GetConfigValue("FTP_USER", "PartDBCfg.py"));
            string ftp_key = (string)(localMethod.GetConfigValue("FTP_KEY", "PartDBCfg.py"));
            var saciftp = new FtpOperationAsync(ftp_addr,ftp_user,ftp_key);
            Action<string,string> syncDB = delegate (string filename, string folderpath)
            {
              

             

                if (!filenameQueue.Contains(filename))
                {
                    //Find the same name file
                    var filter = new BsonDocument("FileName", filename);
                    var check = model.collection.Find(filter);

                    if (check.Count() == 0)
                    {
                        string rev = "";
                        string extension = "";
                        int mindex = filename.LastIndexOf('.');
                        if (mindex >= 2)
                        {
                            rev = filename.Substring(mindex - 2, 2);
                            extension= filename.Substring(mindex);
                 
                        }
                        BsonDocument bd = new BsonDocument {
                    { "FileName",filename },
                    { "Rev",rev },
                    { "Extention",extension},
                     { "FilePath", folderpath+filename },
                      { "InsertDate",DateTime.Now.ToShortDateString()}
                  };
                        model.collection.InsertOne(bd);
                    }
                    else

                    {
                        //检查地址是否存在，若不存在，则更新记录
                        var item = check.First();
                        if (!File.Exists(item["FilePath"].AsString))
                        {
                            var updatestr = new BsonDocument {
                           { "$set",
                           new BsonDocument { { "FilePath", folderpath+filename },
                           { "InsertDate", DateTime.Now.ToShortDateString() }}

                          } };
                            model.collection.UpdateOne(filter, updatestr);
                        }
                    }
                    filenameQueue.Add(filename);
                }
            };
            try
            {
                return await saciftp.WalktreeFTP(path, syncDB);
            }
            catch
            {
                return false;
            }
          
            

        }






    }
}
