using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Bson;
using MongoDB.Driver;

namespace GoumangToolKit.NET4._6
{
    public class PartDB
    {

        MongoMethod model;


       public   PartDB( )
        {
            string client = (string)(localMethod.GetConfigValue("MONGO_URI", "PartDBCfg.py"));
            string database = (string)(localMethod.GetConfigValue("MONGO_DATABASE", "PartDBCfg.py"));
            string collection = (string)(localMethod.GetConfigValue("PART_MONGO_COLNAME", "PartDBCfg.py"));
            model = new MongoMethod(client, database, collection);

        }


      public async Task<bool> UpdatePartDB(IEnumerable<string> PathList)
        {
            Action<FileInfo> syncDB = delegate (FileInfo fi)
              {

                  //Find the same name file
                  var filter = new BsonDocument("FileName", fi.Name);
                  var check =  model.collection.Find(filter);
                  if (check.Count()==0)
                  {
                      int mindex = fi.Name.IndexOf('.');
                      BsonDocument bd = new BsonDocument {
                    { "FileName",fi.Name },
                    { "Rev",fi.Name.Substring(mindex - 2, 2) },
                    { "Extention",fi.Extension},
                     { "FilePath",fi.FullName },
                      { "InsertDate",DateTime.Now.ToShortDateString()}
                  };
                      model.collection.InsertOne(bd);
                  }
              };
            try
            {

              await fileListsAsync.WalkTreeAsync(PathList, syncDB);
                return true;
            }
            catch
            {
                return false;
            }


        }








    }
}
