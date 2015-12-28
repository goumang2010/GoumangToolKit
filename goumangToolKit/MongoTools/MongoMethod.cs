using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GoumangToolKit
{
    public class MongoMethod4
    {
         public MongoDatabase database;
         public MongoClient client;
        public MongoServer server;
         public   MongoCollection<BsonDocument> collection;

        public MongoMethod4(string ConnectStr,string DBNname,string CollectionName )
        {
          //  var credentials =MongoCredential.CreatePlainCredential("SACI", "autorivet", ("222222"));
            var cc = new MongoUrl(ConnectStr);
            client = new MongoClient(cc);
            server = client.GetServer();
            database = server.GetDatabase(DBNname);
            collection = database.GetCollection<BsonDocument>(CollectionName);

           
        }






    }
}
