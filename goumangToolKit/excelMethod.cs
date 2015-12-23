using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace mysqlsolution
{
   public class excelMethod
    {


        public static Worksheet SaveDataTableToExcel(System.Data.DataTable excelTable,string filepath="")
        {
            //creatDir();
            Microsoft.Office.Interop.Excel.Application app =
                new Microsoft.Office.Interop.Excel.ApplicationClass();


            app.Visible = true;
            Workbook wBook = app.Workbooks.Add(true);

            Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
            if (excelTable.Rows.Count > 0)
            {
                int row = 0;
                row = excelTable.Rows.Count;
                int col = excelTable.Columns.Count;
                for (int i = 0; i < row; i++)
                {
                    for (int j = 0; j < col; j++)
                    {
                        string str = excelTable.Rows[i][j].ToString();
                        wSheet.Cells[i + 2, j + 1] = str;
                    }
                }
            }

            int size = excelTable.Columns.Count;
            for (int i = 0; i < size; i++)
            {
                wSheet.Cells[1, 1 + i] = excelTable.Columns[i].ColumnName;
            }
            wSheet.Columns.AutoFit();
            if(filepath!="")
            {
                //设置禁止弹出保存和覆盖的询问提示框 
            app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;

                //保存工作簿 

                //保存excel文件 
             wBook.SaveAs(filepath);
          //  app.SaveWorkspace(filepath);
              app.Quit();
              app = null;
              return null;
            }
            else
            {
               
                return wSheet;
            }

          

        }

        //public static Worksheet SaveDictionaryToExcel(Dictionary<string, object> excelTable, string filepath = "")
        //{
        //    //creatDir();
        //    Microsoft.Office.Interop.Excel.Application app =
        //        new Microsoft.Office.Interop.Excel.ApplicationClass();


        //    app.Visible = true;
        //    Workbook wBook = app.Workbooks.Add(true);

        //    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
           
        //    if (excelTable.Count > 0)
        //    {
               
        //        int row = excelTable.Count;
        //     //   int col = 2;
        //        for (int i = 0; i < row; i++)
        //        {
        //            wSheet.Cells[i + 2, 0] = i + 1;
        //            wSheet.Cells[i + 2, 1] = excelTable.Keys.ElementAt(i);
        //            wSheet.Cells[i + 2, 2] = excelTable.Values.ElementAt(i);
        //        }
        //    }

            
        //    wSheet.Columns.AutoFit();
        //    if (filepath != "")
        //    {
        //        //设置禁止弹出保存和覆盖的询问提示框 
        //        app.DisplayAlerts = false;
        //        app.AlertBeforeOverwriting = false;

        //        //保存工作簿 

        //        //保存excel文件 
        //        wBook.SaveAs(filepath);
        //        //  app.SaveWorkspace(filepath);
        //        app.Quit();
        //        app = null;
        //        return null;
        //    }
        //    else
        //    {

        //        return wSheet;
        //    }



        //}
        //public static void SaveListToExcel(string[] title,List<string> content)
        //{
        //    if(content.Count()>0)
        //    {

        //        int rowqty = content.Count();
        //        int colqty = title.Count();
        //    //creatDir();
        //    Microsoft.Office.Interop.Excel.Application app =
        //        new Microsoft.Office.Interop.Excel.ApplicationClass();


        //    app.Visible = true;
        //    Workbook wBook = app.Workbooks.Add(true);

        //    Worksheet wSheet = wBook.Worksheets[1] as Worksheet;

        //    for (int j = 0; j < colqty; j++)
        //    {


        //        wSheet.Cells[1, j + 1] = title[j];

        //    }
        //        for (int i = 0; i < rowqty; i++)
        //        {
        //            string[] onerow = content.ElementAt(i).Split(',');
        //          for (int j = 0; j < colqty; j++)
        //        {
                 
        
        //                wSheet.Cells[i + 2, j + 1] = onerow[j];

        //           }

                   
        //        }
          

           
        //    //设置禁止弹出保存和覆盖的询问提示框 
        //    //   app.DisplayAlerts = false;
        //    //   app.AlertBeforeOverwriting = false;

        //    //保存工作簿 

        //    //保存excel文件 
        //    //wBook.Save();
        //    //   app.SaveWorkspace(filePath);
        //    //app.Quit();
        //    //  app = null;
        //    wSheet.Columns.AutoFit();
        //    //return wSheet;
        //    }
        //}

        public static bool SaveDataTableToExcel_offset(System.Data.DataTable excelTable, System.Data.DataTable excelTable2)
        {
            //creatDir();
            Microsoft.Office.Interop.Excel.Application app =
                new Microsoft.Office.Interop.Excel.ApplicationClass();


            app.Visible = true;
            Workbook wBook = app.Workbooks.Add(true);

            Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
            int row = 0;
            row = excelTable.Rows.Count;
            int col = excelTable.Columns.Count;
            if (excelTable.Rows.Count > 0)
            {
                
                for (int i = 0; i < row; i++)
                {
                    for (int j = 0; j < col; j++)
                    {
                        string str = excelTable.Rows[i][j].ToString();
                        wSheet.Cells[i + 2, j + 1] = str;
                    }
                }
            }
            int row2 = 0;
            row2 = excelTable.Rows.Count;
            int col2 = excelTable.Columns.Count;
            if (excelTable2.Rows.Count > 0)
            {
                
                for (int i = 0; i < row2; i++)
                {
                    for (int j = 0; j < col2; j++)
                    {
                        string str = excelTable2.Rows[i][j+col].ToString();
                        wSheet.Cells[i + 2, j + 1+col] = str;
                    }
                }
            }








            
            for (int i = 0; i < col; i++)
            {
                wSheet.Cells[1, 1 + i] = excelTable.Columns[i].ColumnName;
            }
            for (int i = 0; i < col2; i++)
            {
                wSheet.Cells[1, 1 + i + col] = excelTable2.Columns[i].ColumnName;
            }
            //设置禁止弹出保存和覆盖的询问提示框 
            //   app.DisplayAlerts = false;
            //   app.AlertBeforeOverwriting = false;

            //保存工作簿 

            //保存excel文件 
            //wBook.Save();
            //   app.SaveWorkspace(filePath);
            //app.Quit();
            //  app = null;
            return true;



        }
        public static bool SaveDataTableToExcelmulti(List<System.Data.DataTable> excelTable, List<string> abcname,string savepath="")
        {
            //creatDir(savepath);
            Microsoft.Office.Interop.Excel.Application app =
                new Microsoft.Office.Interop.Excel.ApplicationClass();
            app.DisplayAlerts = false;
            app.AlertBeforeOverwriting = false;
            if (abcname.Count==1)
            {
                app.Visible = true;
            }
            else
            {
                app.Visible = false;
            }
            
           
            /*        try
                    {
                        wBook.SaveAs(filePath);
                    }

                    catch (Exception err)
                    {


                        /*MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示信息",
             MessageBoxButtons.OK, MessageBoxIcon.Information);
                         * 
                


                        app.Quit();
                        app = null;
                        MessageBox.Show("该文件已打开,请关闭后重试");
                        return false;
                    }
        **/
            int count = excelTable.Count();
            for (int k = 0; k < count; k++)
            {
                Workbook wBook = app.Workbooks.Add(true);
               // wBook..Add();
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;


                System.Data.DataTable excelsheet = excelTable[k];
                if (excelsheet.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelsheet.Rows.Count;
                    int col = excelsheet.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            string str = excelsheet.Rows[i][j].ToString();
                            //wSheet.Cells[i + 2, j + 1] = str;
                            wSheet.Cells[j + 2, i + 2] = str;
                            Range mycell=(Range)wSheet.Cells[j + 2, i + 2];
                                
                               mycell.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.Black.ToArgb());

                        }
                    }

                    
                }

                int size = excelsheet.Columns.Count;
                for (int i = 0; i < size; i++)
                {
                    wSheet.Cells[2 + i,1 ] = excelsheet.Columns[i].ColumnName;

                    Range mycell = (Range)wSheet.Cells[2 + i, 1];

                    mycell.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.Black.ToArgb());

 

                }
                wSheet.Cells[1, 1] = abcname[k];

                Range allColumn = wSheet.Columns;
                Range allColumn3 = wSheet.get_Range(wSheet.Cells[2, 1], wSheet.Cells[excelsheet.Columns.Count+1, excelsheet.Rows.Count + 1]);
                  allColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                  allColumn.ColumnWidth = 15;
                  allColumn.WrapText = true;
                  Range allColumn2 = (Range)wSheet.Columns[1];
                  allColumn2.ColumnWidth = 18.5;
                  allColumn2.RowHeight = 27;
                  Range allColumn4= (Range)wSheet.Rows[5];
                  allColumn4.RowHeight = 50;
                  allColumn.AutoFit();
                  wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells[1, excelsheet.Rows.Count + 1]).Merge();
                  //allColumn3.Borders.LineStyle=1;
               allColumn3.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.Black.ToArgb());
               // wSheet.Name = abcname[k];
                  
                if(savepath!="")
                {
                    wBook.SaveAs(savepath + "\\" + abcname[k]);
                }
             

            }
            //设置禁止弹出保存和覆盖的询问提示框 
 

            //保存工作簿 

            //保存excel文件 
            //   wBook.Save();
            //   app.SaveWorkspace(filePath);
            if (abcname.Count != 1)
            {
                app.Quit();
                app = null;
            }

            return true;

        }
        public static bool SaveDataTableToExcelM(List<System.Data.DataTable> excelTable, List<string> abcname)
        {
            creatDir();
            Microsoft.Office.Interop.Excel.Application app =
                new Microsoft.Office.Interop.Excel.ApplicationClass();
           

            app.Visible = true;
            Workbook wBook = app.Workbooks.Add(true);
            
            /*        try
                    {
                        wBook.SaveAs(filePath);
                    }

                    catch (Exception err)
                    {


                        /*MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示信息",
             MessageBoxButtons.OK, MessageBoxIcon.Information);
                         * 
                


                        app.Quit();
                        app = null;
                        MessageBox.Show("该文件已打开,请关闭后重试");
                        return false;
                    }
        **/
            int count = excelTable.Count();
            for (int k = 0; k < count; k++)
            {

                wBook.Worksheets.Add();
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;


                System.Data.DataTable excelsheet = excelTable[k];
                if (excelsheet.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelsheet.Rows.Count;
                    int col = excelsheet.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            string str = excelsheet.Rows[i][j].ToString();
                            wSheet.Cells[i + 2, j + 1] = str;
                        }
                    }
                }

                int size = excelsheet.Columns.Count;
                for (int i = 0; i < size; i++)
                {
                    wSheet.Cells[1, 1 + i] = excelsheet.Columns[i].ColumnName;
                }
                wSheet.Name = abcname[k];
            }
            //设置禁止弹出保存和覆盖的询问提示框 
            //   app.DisplayAlerts = false;
            //   app.AlertBeforeOverwriting = false;

            //保存工作簿 

            //保存excel文件 
            //   wBook.Save();
            //   app.SaveWorkspace(filePath);
            //app.Quit();
            //  app = null;
            return true;

        }
        private static void creatDir()
        {

            if (!Directory.Exists("D:\\绩效考核\\"))
            {
                Directory.CreateDirectory("D:\\绩效考核\\");
            }


        }
       public static void creatDir(string savepath)
        {

            if (!Directory.Exists(savepath))
            {
                Directory.CreateDirectory(savepath);
            }


        }

        //加载Excel 
        public static System.Data.DataTable LoadDataFromExcel(string filePath)
        {


            try
            {
                string strConn;
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=False;IMEX=1'";
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                String sql = "SELECT * FROM  [Sheet1$]";//可是更改Sheet名称，比如sheet2，等等 

                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                System.Data.DataTable OleDsExcle = new System.Data.DataTable();
                OleDaExcel.Fill(OleDsExcle);
                OleConn.Close();
                return OleDsExcle;
            }
            catch (Exception err)
            {
                MessageBox.Show("读取Excel失败!失败原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }
        public static List<List<string>> LoadDataFromExcelviaapp(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app =
               new Microsoft.Office.Interop.Excel.ApplicationClass();
            List<List<string>> abc = new List<List<string>>();

            app.Visible = false;
            try
            {
                Workbook wBook = app.Workbooks.Open(filePath);
                int sheetqty = wBook.Worksheets.Count;
                for (int k = 1; k <= sheetqty; k++)
                {

                    Worksheet wSheet = wBook.Worksheets[k] as Worksheet;
                    if (wSheet.Name.Contains("组"))
                    {
                        List<string> temptable = new List<string>();
                        int mm = 6;
                        object kkkk = wSheet.Cells[mm, 2];
                        string aaaa = ((Range)kkkk).Text.ToString();
                        while (aaaa != "")
                        {
                            temptable.Add(((Range)wSheet.Cells[mm, 2]).Text.ToString());
                            mm = mm + 1;
                            kkkk = wSheet.Cells[mm, 2];
                            aaaa = ((Range)kkkk).Text.ToString();


                        }

                        abc.Add(temptable);
                    }



                }

                    app.Quit();
                



                return abc;


            }
            catch (Exception err)
            {
                MessageBox.Show("读取Excel失败!失败原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }
    }
}
