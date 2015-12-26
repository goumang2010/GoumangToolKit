using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using System.Data;
using System.IO;
using System.Security.Cryptography;
using Newtonsoft.Json;
using System.Windows.Forms;
using GoumangToolKit;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using Microsoft.VisualBasic;
namespace GoumangToolKit
{
   public static class localMethod
    {
        public static dynamic GetConfigValue(string name, string pypath="configuration.py")
        {
            pypath = @"\\192.168.3.32\softwareTools\Autorivet_team_manage\settings\" + pypath;
            ScriptEngine engine = Python.CreateEngine();
            ScriptScope scope = engine.CreateScope();
           
            if (File.Exists(pypath))
            {
                engine.ExecuteFile(pypath, scope);
                return scope.GetVariable(name);
            }
            else
            {
                return @"\\192.168.3.32\Autorivet\prepare\INFO\";

            }

        }
        public static void UpdateConfigValue(string oldstr,string newstr,string pypath ="CouponCfg.py")
        {
            
            pypath = @"\\192.168.3.32\softwareTools\Autorivet_team_manage\settings\" + pypath;
            backupfile(pypath);
            var lines = ReadLines(pypath);
           var dd= lines.Select(p => p.Replace(oldstr, newstr));
            

            if(dd!=null&&dd.Count()!=0)
            {
                //Copy to list,so it can close the file
                WriteFile(dd.ToList(), pypath);
            }
           

        }
        public static string VBInputBox(string hint="",string title="",string defaultStr="")
        {
          return  Interaction.InputBox(hint, title, defaultStr);
        }



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
        public static void creatDir(string savepath)
        {

            if (!Directory.Exists(savepath))
            {
                Directory.CreateDirectory(savepath);
            }


        }

        public static Form get_Form(string formname)
        {

            foreach (Form Frm in System.Windows.Forms.Application.OpenForms)
            {
                if (Frm.Name == formname)
                {
                    return Frm;

                }
            }
            return null;
        }

        //public static bool WriteLines(IEnumerable<string> content, string fileName)
        //{
        //    using (StreamWriter sw = new StreamWriter(fileName, false))
        //    {
        //        foreach(string pp in content)
        //        {
        //            sw.WriteLine(pp);
        //        }
               
        //    }
        //    return true;
        //}

      public  static IEnumerable<string> ReadLines(string filename)
        {
            return ReadLines(filename, Encoding.UTF8);
        }

      public  static IEnumerable<string> ReadLines(string filename, Encoding encoding)
        {
            return ReadLines(delegate {
                return new StreamReader(filename, encoding);
            });
        }

     public  static IEnumerable<string> ReadLines(Func<TextReader> provider)
        {
            using (TextReader reader = provider())
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    yield return line;
                }
            }
        }


        public static bool WriteFile(this IEnumerable<string> ls,string filepath)
        {
            using (StreamWriter sw = new StreamWriter(filepath,false))
            {
                foreach(string pp in ls)
                {
                    sw.WriteLine(pp);
                }
                
            }
            return true;
        }

        






        public static void SetVisible(this IEnumerable<Control> tt, bool vis)

        {
            foreach (var pp in tt)
            {
                pp.Visible = vis;
            }

        }
        public static void SetPropValue(this IEnumerable<Control> tt,string prop, object value)

        {
            foreach (var pp in tt)
            {
              
                Type type = pp.GetType(); //获取类型
                System.Reflection.PropertyInfo propertyInfo = type.GetProperty(prop);
                propertyInfo.SetValue(pp, value, null); //给对应属性赋值
            }
          
          
               

            }

        public static string listtotext(List<string> abc)
        {

            string newtext = "";

            foreach (string ddd in abc)
            {
                newtext = newtext + ddd + "\r\n";
            }
          
            return newtext;
        }
        public static string GetMD5HashFromFile(string fileName)
     {  
           try  
          {  
              FileStream file = new FileStream(fileName, System.IO.FileMode.Open);  
              MD5 md5 = new MD5CryptoServiceProvider();  
              byte[] retVal = md5.ComputeHash(file);  
              file.Close();  
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < retVal.Length; i++)
                {
                    sb.Append(retVal[i].ToString("x2"));
                   
                }
                return sb.ToString();
            }  
           catch (Exception ex)  
           {  
                throw new Exception("GetMD5HashFromFile() fail,error:" + ex.Message);
               
            }
            return null;
        }



       // public static string ToJson(DataTable dt)
       //{  
            
       //    StringBuilder jsonString = new StringBuilder();  
       //jsonString.Append("[");  
       //  DataRowCollection drc = dt.Rows;  
       //    for (int i = 0; i<drc.Count; i++)  
       //    {  
       //       jsonString.Append("{");  
       //      for (int j = 0; j<dt.Columns.Count; j++)  
       //       {  
       //             string strKey = dt.Columns[j].ColumnName;  
       //             string strValue = drc[i][j].ToString();  
       //             Type type = dt.Columns[j].DataType;  
       //           jsonString.Append("\"" + strKey + "\":");  
       //           strValue = StringFormat(strValue, type);  
       //            if (j<dt.Columns.Count - 1)  
       //           {  
       //                 jsonString.Append(strValue + ",");  
       //           }  
       //            else  
       //             {  
       //                jsonString.Append(strValue);  
       //            }  
       //         }  
       //         jsonString.Append("},");  
       //   }  
       //     jsonString.Remove(jsonString.Length - 1, 1);  
       //     jsonString.Append("]");  
       //    if (jsonString.Length == 1)  
       //    {  
       //         return "[]";  
       //  }  
       //     return jsonString.ToString();  
       // }  


        public static string ToJson(Dictionary<string,string> dic)
        {
            StringWriter sw = new StringWriter();
            JsonWriter writer = new JsonTextWriter(sw);

            writer.WriteStartObject();
            
            foreach (var dicPair in dic)
            {
                writer.WritePropertyName(dicPair.Key);

                writer.WriteValue(dicPair.Value);
            }


            writer.WriteEndObject();
            writer.Flush();

            string jsonText = sw.GetStringBuilder().ToString();
         return  jsonText;


        }






        public static string GetProcessUserName(int pID)
        {


            string text1 = null;


            SelectQuery query1 =
                new SelectQuery("Select * from Win32_Process WHERE processID=" + pID);
            ManagementObjectSearcher searcher1 = new ManagementObjectSearcher(query1);


            try
            {
                foreach (ManagementObject disk in searcher1.Get())
                {
                    ManagementBaseObject inPar = null;
                    ManagementBaseObject outPar = null;


                    inPar = disk.GetMethodParameters("GetOwner");


                    outPar = disk.InvokeMethod("GetOwner", inPar, null);


                    text1 = outPar["User"].ToString();
                    break;
                }
            }
            catch
            {
                text1 = "SYSTEM";
            }


            return text1;
        }
      

       


       public static string skin_to_drawing(string skinname)
        {
          var temp2rear = skinname.Substring(10, 3);
          var temp2head = skinname.Substring(0, 9);

          var temp2rearint = Convert.ToInt32(temp2rear);

        var dt=  DbHelperSQL.Query("select 蒙皮号,图号 from 产品列表 where 蒙皮号 like '" + temp2head + "%'").Tables[0];
        
      foreach(DataRow pp in dt.Rows)
      {
          int aa = Convert.ToInt32(pp["蒙皮号"].ToString().Split('-')[1]);
          int nn = Math.Abs(aa - temp2rearint);
          if(nn%2==0)
          {
              return pp["图号"].ToString().Replace("-001", "");
          }
      }
      return "";

        }

     




    }
}
