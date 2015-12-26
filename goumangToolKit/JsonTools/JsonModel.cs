using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace GoumangToolKit
{
    class JsonModel
    {
    }

   public class historyJsonModel
    {
       // public string filepath { get; set; }
        public string writer { get; set; }
        public string date { get; set; }
        public string time { get; set; }
        public string operation { get; set; }

        public override string ToString()
        {
            string outputstring = "Writer:"+writer+"\r\nOperation:"+operation+"\r\nDate:"+date+"\r\nTime:"+time+ "\r\n\r\n";
            return outputstring;
        }

    }

    public static class jsonMethod
    {
        public static List<historyJsonModel> iniFromStr( string str)
        {
            // hmls = new List<historyJsonModel>();
            JsonSerializer serializer = new JsonSerializer();
            StringReader sr = new StringReader(str);
            //若是直接赋值，则不能使用静态拓展方法
           return (List<historyJsonModel>)serializer.Deserialize(new JsonTextReader(sr), typeof(List<historyJsonModel>))??new List<historyJsonModel>();
           // hmls = (p1 ?? new List<historyJsonModel>());
           
        }

        public static string JsonToString(this List<historyJsonModel> hmls)
        {
            string output = "";
            foreach(var pp in hmls)
            {
                output+= pp.ToString();

            }

            return output;
        }
        public static string ToJson(this List<historyJsonModel> p1)
        {
            JsonSerializer serializer = new JsonSerializer();
            StringWriter sb = new StringWriter();
            serializer.Serialize(new JsonTextWriter(sb), p1);


            return sb.ToString();
        }
      



    }










}
