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
   public static class FileOperation
    {
   

   
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

        



    }
}
