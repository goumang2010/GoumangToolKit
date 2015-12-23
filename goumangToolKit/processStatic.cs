using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using OFFICE_Method;

namespace mysqlsolution
{
   public class processStatic
    {
       public Dictionary<string, object> Points_ = new Dictionary<string, object>();


        //  Private Points_new As New Dictionary(Of String, Integer)

            



      public void Add(int qty, string processtype)
      {

          if (qty > 0)
          {
              if (InTheList(processtype))
              {
                  //数量+qty
                  Points_[processtype] =(int)Points_[processtype]+ qty;

              }
              else
              {
                  Points_.Add(processtype, qty);

              }



          }
      }


      public void Add(object qty, string processtype)
      {

         
              if (InTheList(processtype))
              {
                  //数量+qty
                  Points_[processtype] = qty;

              }
              else
              {
                  Points_.Add(processtype, qty);

              }



          
      }



       //返回该类型数量
       public object Item(string processtype)
       {
           return Points_[processtype];
       }


       //返回第几个数量

              public object Item2(int index)
       {
           return Points_.Values.ElementAt(index-1);
       }


       //返回 加工类型
                     public string Key(int index)
       {
           return Points_.Keys.ElementAt(index-1);
       }



       public void Add(processStatic mm)
       {
          foreach(var kk in mm.Points_)
          {
              Add(kk.Value, kk.Key);


          }

       }
    //       Public Function count() As Integer
    //    count = Points_.Count
    //End Function


       public int count()
       {
           return Points_.Count();
       }



       public bool InTheList(string processtype)
       {
           return Points_.ContainsKey(processtype);
       }



    //Public Function Remove(Index As Integer)

    //    If (Index > 0 And Index <= Points_.Count) Then
    //        Points_.Remove(Index)
    //        Points_index.Remove(Index)
    //    End If

    //End Function




       public void Remove(int Index)
       {
           Points_.Remove(Key(Index));
       }


             public void RemoveName(string  processtype)
       {
           Points_.Remove(processtype);

       }



             public void RemoveAll()
             {
                 Points_.Clear();

             }

    //      Public Sub report()
    

    //    Dim excelobject As Object
    //    Dim wb As Object

    //    excelobject = CreateObject("excel.application") '启动Excel程序
    //    excelobject.Visible = True   '可见
    //    wb = excelobject.Workbooks.Add()



    //    Dim m As Integer
    //    Dim n As Integer
    //    Dim p As Integer



    //    For m = 1 To count()

    //        Dim changestr As String


    //        changestr = Key(m)
    //        Dim kkk As String()
    //        kkk = Strings.Split(changestr, " - ")
    //        Dim cnt As Integer = kkk.Count
    //        For i As Integer = 0 To cnt - 1

    //            wb.Sheets(1).Cells(m + 4, i + 1).Value = kkk(i)
    //        Next
    //        'wb.Sheets(1).Cells(m + 4, 1).Value = Strings.Split(changestr, " - ")(0)
    //        'wb.Sheets(1).Cells(m + 4, 2).Value = Strings.Split(changestr, " - ")(1)
    //        '取得该种类型的个数
    //        wb.Sheets(1).Cells(m + 4, cnt + 1).Value = Item2(m)

    //    Next


    //    wb.Sheets(1).Cells.EntireColumn.AutoFit()

    //End Sub

       public void report()
             {



            OFFICE_Method.excelMethod.SaveDataTableToExcel(output_dt());
           // excelMethod.;
           //      excelMethod.SaveDataToExcel(kkk);

             }
        public DataTable output_dt()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("QTY", typeof(string));
            dt.Columns.Add("DESP", typeof(string));

            foreach (var item in Points_)
            {
                DataRow dr = dt.NewRow();
                dr["QTY"] = item.Key;
                dr["DESP"] = item.Value;
                dt.Rows.Add(dr);

            }
            return dt;

        }

        public int SearchIndex(string  processtype)
       {
           if (InTheList(processtype))
           {
              return Points_.Keys.ToList().FindIndex(p => p == processtype) + 1;

           }
           else
           {
               return -1;
           }


       }








    }
}
