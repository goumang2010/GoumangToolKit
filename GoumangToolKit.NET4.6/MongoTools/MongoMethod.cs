using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GoumangToolKit
{
    public class MongoMethod
    {
         public IMongoDatabase database;
         public IMongoClient client;
         public   IMongoCollection<BsonDocument> collection;

        public MongoMethod(string ConnectStr,string DBNname,string CollectionName )
        {
            var client = new MongoClient(ConnectStr);
            var database = client.GetDatabase(DBNname);
            collection = database.GetCollection<BsonDocument>(CollectionName);


        }

        public async Task<IEnumerable<BsonDocument>> FetchTextData(string filterstr)
        {
            string regEx = "-";
            var array = Regex.Split(filterstr, regEx, RegexOptions.IgnoreCase);
            string querystr = "";
            if (array.Count() == 1)
            {
                querystr = filterstr;
            }
            else
            {
                foreach (var pp in array)
                {
                    querystr = querystr + "\"" + pp + "\"";

                }
            }



            var filter = Builders<BsonDocument>.Filter.Eq("$text", new BsonDocument { { "$search", querystr } });
            List<string> dd = new List<string>();
            return await collection.Find(filter).ToListAsync();
        }




    }
}
