using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace mysqlsolution
{
   public class AutorivetDB
   {

       public static string trimprod(string productname)
   {
       productname = productname.Replace("-001", "");
       productname = productname.Replace("process", "");
       return productname;
   }
       //操作工艺文件表

          #region paperwork




       public static string getpath(string filename)
       {
           string path = DbHelperSQL.getlist("select 文件地址 from paperWork where 文件名='" + filename + "'").First();
           return path;


       }







       private static string removepath(string sqlstr)
       {

           return sqlstr.Replace(",文件地址 as 地址", "");

       }


       #region paperwork


       public static DataTable paperwork_view(int i = 0)
       {
           string sqlstr = "select 文件类型 as 类型,文件名,文件编号,文件名称,状态,位置,版次,关联产品,关联架次,编制日期,修改日期,备注,文件地址 as 地址 from paperWork order by 编制日期 desc";
           //已建立视图
           // string sqlstr = "select * from paperwork_view;";
           if (i == 1)
           {
               sqlstr = removepath(sqlstr);
           }

           return DbHelperSQL.Query(sqlstr).Tables[0];








       }

       public static DataTable ppworkviewrf(int k, int i = 0)
       {
           //用于判断筛选的文件类型
           switch (k)
           {


               case 1:
                   //AO
                   return ppworkviewrf("AO", 2, i);

               case 2:
                   //补铆_AAO
                   return ppworkviewrf("补铆_AAO", 2, i);

               case 3:
                   //RNC_AAO
                   return ppworkviewrf("RNC_AAO", 2, i);

               case 4:
                   //COS
                   return ppworkviewrf("COS", 2, i);

               case 5:
                   //技术单
                   return ppworkviewrf("TS", 2, i);

               case 6:
                   //大纲索引
                   return ppworkviewrf("INDEX", 2, i);

               default:
                   //显示全部
                   return paperwork_view(i);

           }
       }

       public static DataTable ppworkviewrf(string filterstr, int k, int i = 0)
       {
           string sqlstr = "";



           switch (k)
           {
               case 1:
                   //文件编号
                   sqlstr = "select 文件类型 as 类型,文件名,文件编号,文件名称,版次,关联产品,关联架次,编制日期,修改日期,备注,文件地址 as 地址 from paperWork where 文件编号='" + filterstr + "' order by 编制日期 desc";

                   break;
               case 2:
                   //文件类型
                   sqlstr = "select 文件类型 as 类型,文件名,文件编号,文件名称,版次,关联产品,关联架次,编制日期,修改日期,备注,文件地址 as 地址 from paperWork where 文件类型='" + filterstr + "' order by 编制日期 desc";
                   break;
           }

           if (i == 1)
           {
               sqlstr = removepath(sqlstr);
           }
           return DbHelperSQL.Query(sqlstr).Tables[0];
       }

       #endregion

























     
          #endregion

        #region otherPaperWork

        public static void otherpaperwork_table()
        {
            DbHelperSQL.ExecuteSql("Create table if not exists otherPaperWork(文件名 varchar(100),文件类型 varchar(100),文件编号 varchar(100) NOT NULL PRIMARY KEY,文件名称 varchar(100),版次 varchar(20),编制日期 date default null ,修改日期 date default null ,文件地址 varchar(100),状态 varchar(10),位置 varchar(10),关联产品 varchar(100),关联架次 varchar(100),相关文件 varchar(100),文件格式 varchar(100),工作包 varchar(100),备注 text);");

          
        }

        public static DataTable otherpw(string pkg)
        {
           return DbHelperSQL.Query("select 文件编号,文件名称,关联产品,文件格式 as 格式,修改日期 from otherpaperwork where 工作包='"+pkg+"' order by 文件编号" ).Tables[0];
        }

        //public static string getopwpath()
        //{
        //    return DbHelperSQL.i("select 文件编号,文件名称,关联产品,文件格式,修改日期,工作包 from otherpaperwork ").Tables[0];
        //}
        #endregion


        //操作Process表
        #region Process_op
       public static List<string> processitemlist(string productname,string item)
       {
           string prodname = trimprod(productname) + "Process";

           return DbHelperSQL.getlist("select " + item + " from " + prodname + " group by " + item);



       }

        public static DataTable getparatable(string productname)
        {
            string prodname = trimprod(productname) + "Process";

            return DbHelperSQL.Query("select ID,UUID,加工位置location,紧固件名称Fastener_Name,紧固件数量Fastener_Qty,钻头Drill,下铆头Lower_Anvil,上铆头Upper_Anvil,胶嘴Sealant_Tip,试片Coupon_used,参数号Process_NO,锪窝深度Countersink_depth,钻头转速Speed_of_drill,给进速率Feed_speed,夹紧力Clamp_force,夹紧释放力Clamp_relief_force,墩铆力Upset_force,墩铆位置Upset_position,注胶压力Seal_pres,注胶时间Seal_time from " + prodname + " order by ID").Tables[0];



        }




        #endregion


        //操作TVA
        #region TVA_op

        //更新location及strno



        public static void updateTVAlocation(string proname)

        {

            proname = trimprod(proname);


            var cc = DbHelperSQL.getlist("select XR,count(*)  as xqty from " + proname + " group by XR order by xqty desc");

            if (cc.Count != 0)
            {





                int chosno;

                chosno = System.Convert.ToInt32(cc.First());
                var seldir = new DataTable();
                var sqlstr33 = new StringBuilder();
                sqlstr33.Append("select YR,ZR from " + proname);
                sqlstr33.Append(string.Format(" where XR={0} order by STRNO", chosno.ToString()));

                seldir = DbHelperSQL.Query(sqlstr33.ToString()).Tables[0];


                var ygap = System.Convert.ToInt32(seldir.Rows[0][0].ToString()) - System.Convert.ToInt32(seldir.Rows[seldir.Rows.Count - 1][0].ToString());
                var zgap = System.Convert.ToInt32(seldir.Rows[0][1].ToString()) - System.Convert.ToInt32(seldir.Rows[seldir.Rows.Count - 1][1].ToString());



                var len = System.Math.Sqrt(ygap* ygap + zgap * zgap);
                var py = ygap / len;
                var pz = zgap / len;

                if (System.Math.Abs(py) > 0.9)
                {
                    py = 1;
                    pz = 0;
                }

                else
                {
                   if (System.Math.Abs(pz) > 0.9)
                {
                    pz = 1;
                    py = 0;
                }


                }


                var strSqlnamepr = new StringBuilder();
                strSqlnamepr.Append(string.Format("UPDATE {0} set ", proname));

            strSqlnamepr.Append(string.Format("location='',strno=round(SQRT(POW({0}*(YR),2)+POW({1}*(ZR),2)))*YR/ABS(YR) where location <>'WIN' and location <>'DOU'", py, pz));
                DbHelperSQL.ExecuteSql(strSqlnamepr.ToString());

               
                var locationlist= new List<string>();
                //得到确定是长桁的长桁ID
                locationlist = DbHelperSQL.getlist("select STRNO from(select STRNO,count(*) as qty from " + proname + " group by STRNO) = org where qty>20");
                
                foreach (var strno in locationlist)
               {


                    var saa = new StringBuilder();

                    saa.Append(string.Format("UPDATE {0} set ", proname));

                    saa.Append(string.Format("location='STR' where STRNO='{0}'", strno));
                    DbHelperSQL.ExecuteSql(saa.ToString());

                    }




                //进行二次筛选，防止四舍五入造成的误差

                //首先取得还没有标识的组

                var otherlist = DbHelperSQL.getlist("select STRNO from " + proname + " where location='' and strno is not null");
                //对该组内每个元素进行遍历,筛查是否有相近的元素
                var buchong = new List<string>();

               for(int i = 0;i< (otherlist.Count - 1);i++)
                {





                    var pp_strno = otherlist[i];
                   foreach (string strno in locationlist)
                    {
                        if(Math.Abs(System.Convert.ToInt32(pp_strno) - System.Convert.ToInt32(strno)) < 2)
                        {
                            buchong.Add(pp_strno + "_" + strno);

                            break;
                        }
                    }
                       

             }

                foreach (string item in buchong)
                {

                    var kk = item.Split('_');

                    var strSqlname = new StringBuilder();

                    strSqlname.Append(string.Format("UPDATE {0} set ", proname));
                    strSqlname.Append(string.Format("location='STR',STRNO='{0}' where STRNO='{1}'", kk[1], kk[0]));
                    DbHelperSQL.ExecuteSql(strSqlname.ToString());
                }

  
               //未识别成长桁或窗框的则识别为FRM
                var strSqlname2 = new StringBuilder();
                strSqlname2.Append(string.Format("UPDATE {0} set ", proname));
                strSqlname2.Append("location='FRM',STRNO=XR where location=''");
                DbHelperSQL.ExecuteSql(strSqlname2.ToString());

                //重制框的标识

                var frmlist = DbHelperSQL.getlist("select XR from(select XR,count(*) as qty from " + proname + " where location='FRM' group by XR ) as org where qty>10");

                var frmpplist = DbHelperSQL.getlist("select XR from " + proname + " where location='FRM'");

              foreach(string pp in frmpplist)
                        {

             

                    if (!frmlist.Contains(pp))
                    {



                        foreach (string dd in frmlist)
                        {

                      
                            if (Math.Abs(System.Convert.ToInt32(pp) - System.Convert.ToInt32(dd)) <2)
                            {





                                var strSqlname = new StringBuilder();
                                strSqlname.Append(string.Format("UPDATE {0} set ", proname));

                                strSqlname.Append(string.Format("STRNO='{0}' where location='FRM'and STRNO='{1}'", dd, pp));
                                DbHelperSQL.ExecuteSql(strSqlname.ToString());


                                break;
                            }

                        }



                  

                    }

                }

                //经过以上的处理过程，可以实现对单曲壁板的重新规划



                //var updatelist = new List<string>();

                //updatelist = (DbHelperSQL.getlist("select UUID from " + proname + "process_backup"));

                DbHelperSQL.ExecuteSql("delete from " + proname + "process_backup");


                //if (updatelist.Count() == 0)
                //{







                    var loweranvil = "(case when t1.location='STR' and t2.Cycle_Type='M60' then '401C260' when t1.location<>'STR' and t2.Cycle_Type='M60' then '401C263' else '401C265' end) as loweranvil";
                    var tiqu1 = "(select location,FastenerName,count(*) as qty from " + proname + " WHERE ProcessType like '%INSTALLED BY%' group by location,FastenerName  order by location desc,FastenerName) t1";
                    var tiqu = "select CONCAT(t1.location,t1.FastenerName),t1.location,t1.FastenerName,t1.qty,t2.Drill," + loweranvil + ",t2.Upper_Anvil,t2.Tips from (" + tiqu1 + " left join 紧固件列表 t2 on t1.FastenerName=t2.Fasteners) order by t2.Tips,loweranvil";
                    DbHelperSQL.ExecuteSql("insert into " + proname + "process_backup(uuid,加工位置location,紧固件名称Fastener_Name,紧固件数量Fastener_Qty,钻头Drill,下铆头Lower_Anvil,上铆头Upper_Anvil,胶嘴Sealant_Tip) " + tiqu);
                    DbHelperSQL.ExecuteSql("update " + proname + "process_backup tt inner join (SELECT t.UUID,t.胶嘴Sealant_Tip as Tips,t.下铆头Lower_Anvil as loweranvil ,@rownum := @rownum + 1 AS rank FROM " + proname + "process_backup t,(SELECT @rownum := 0) r order by Tips,loweranvil) kk on tt.UUID=kk.UUID  set tt.ID=kk.Rank");

                //        }

                //else
                //{



                //    var loweranvil = "(case when t1.location='STR' and t2.Cycle_Type='M60' then '401C260' when t1.location<>'STR' and t2.Cycle_Type='M60' then '401C263' else '401C265' end) as loweranvil";
                //    var tiqu1 = "(select location,FastenerName,count(*) as qty from " + proname + " WHERE ProcessType like '%INSTALLED BY%' group by location,FastenerName  order by location desc,FastenerName) t1";
                //    var tiqu = "select CONCAT(t1.location,t1.FastenerName) as uuid,t1.location,t1.FastenerName,t1.qty,t2.Drill," + loweranvil + ",t2.Upper_Anvil,t2.Tips from (" + tiqu1 + " left join 紧固件列表 t2 on t1.FastenerName=t2.Fasteners) order by t2.Tips,loweranvil";
                //    DbHelperSQL.ExecuteSql("insert ignore into " + proname + "process_backup(uuid,加工位置location,紧固件名称Fastener_Name,紧固件数量Fastener_Qty,钻头Drill,下铆头Lower_Anvil,上铆头Upper_Anvil,胶嘴Sealant_Tip) " + tiqu);
                //    DbHelperSQL.ExecuteSql("update " + proname + "process_backup pp inner join (" + tiqu + ") kk on pp.uuid=kk.uuid set pp.加工位置location=kk.location,pp.紧固件名称Fastener_Name=kk.FastenerName,pp.紧固件数量Fastener_Qty=kk.qty,pp.钻头Drill=kk.Drill,pp.下铆头Lower_Anvil=kk.loweranvil,pp.上铆头Upper_Anvil=kk.Upper_Anvil,pp.胶嘴Sealant_Tip=kk.Tips");
                        
                //        //if (CheckBox7.Checked == true)
                //    //{



                //    //    DbHelperSQL.ExecuteSql("update " + proname + "process_backup tt inner join (SELECT t.UUID,t.胶嘴Sealant_Tip = Tips,t.下铆头Lower_Anvil = loweranvil ,@rownum := @rownum + 1 = rank FROM " + proname + "process_backup t,(SELECT @rownum := 0) r order by Tips,loweranvil) kk on tt.UUID=kk.UUID  set tt.ID=kk.Rank")
                //    //  }

                //}





                }

            }



        

        


















       //获取TVA安装的紧固件信息
       
        public static DataTable TVA_fstlist(string productname, string typestr="INSTALLED BY")
        {
            //自动适配参数

            productname = trimprod(productname);

            DataTable fsttemp = DbHelperSQL.Query("select FastenerName,count(*) as qty from " + productname + " WHERE ProcessType like '%" + typestr + "%' group by FastenerName order by FastenerName").Tables[0];
            return fsttemp;
        }
       //TVA中筛选紧固件数量
        public static int fst_qty(string productname, string fsttype, string proctype = "INSTALLED BY")
        {
            //自动适配参数

            productname = trimprod(productname);

            DataTable fsttemp = DbHelperSQL.Query("select count(*) as qty from " + productname + " WHERE FastenerName like '%" + fsttype + "%' and ProcessType like '%" + proctype + "%'").Tables[0];
            return System.Convert.ToInt16( fsttemp.Rows[0][0].ToString());
        }

       #endregion


        //对于产品列表的操作


        #region Product_op


        public static void product_table()
        {
            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.InitialDirectory = "D://";

            fileDialog.Filter = "xls files (*.xls,*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";

            fileDialog.FilterIndex = 1;

            fileDialog.RestoreDirectory = true;

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {

                DbHelperSQL.ExecuteSql("Create table if not exists 产品列表(蒙皮号 varchar(100),图号 varchar(100) PRIMARY KEY,图纸名称 varchar(100),名称 varchar(100),大纲编号 varchar(100),程序编号 varchar(100),状态编号 varchar(100),站位号 varchar(50),下级装配号 varchar(50),图纸版次 varchar(50),AOI编号 varchar(100),预铆编号 varchar(100),AO varchar(100),AOI varchar(100),PACR varchar(100),COS varchar(100),TS varchar(100),AO_INDEX varchar(100),PROGRAM varchar(100),AAO int,生产 int);");
                //DbHelperSQL.ExecuteSql("delete from 产品列表");

                DataTable testExcel =OFFICE_Method.excelMethod.LoadDataFromExcel(fileDialog.FileName);
                DataRow[] shuju;

                shuju = testExcel.Select();
                List<string> creatName = new List<string>();

                foreach (DataRow p in shuju)
                {




                    if (p[0].ToString() != "")
                    {



                        StringBuilder strSqlname = new StringBuilder();

                        strSqlname.Append("REPLACE INTO 产品列表 (");

                        strSqlname.Append("蒙皮号,图号,图纸名称,名称,大纲编号,程序编号,状态编号,站位号,下级装配号,图纸版次,AOI编号,预铆编号");

                        strSqlname.Append(String.Format(") VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", p[0].ToString(), p[1].ToString(), p[2].ToString(), p[3].ToString(), p[4].ToString(), p[5].ToString(), p[6].ToString(), p[7].ToString(), p[8].ToString(), p[9].ToString(), p[10].ToString(), p[11].ToString()));

                        creatName.Add("Create table if not exists " + p[1].ToString().Replace("-001", "") + "(PFname varchar(100) ,FastenerName varchar(100),X double,Y double,Z double,FrameName varchar(100),ProcessType varchar(100),XR INT(20),YR INT(20),ZR INT(20),UUID varchar(100) PRIMARY KEY,STRNO varchar(100),location varchar(100),pointname varchar(100),UUIDP varchar(100));");

                        creatName.Add(strSqlname.ToString());

                    }


                }


                int k = DbHelperSQL.ExecuteSqlTran(creatName);

                MessageBox.Show(String.Format("执行成功,增加 '{0}'条记录", k));



            }
        }
        public static List<string> fullname_list()
        {
            List<string> templist = DbHelperSQL.getlist("select concat(图号,'_',名称,站位号) from 产品列表");
            return templist;
        }
        public static List<string> fullname_list(string namestring)
        {
            List<string> templist = DbHelperSQL.getlist("select "+namestring+" from 产品列表");
            return templist;
        }
        public static DataTable fullname_table(string namestring)
        {
            DataTable templist = DbHelperSQL.Query("select " + namestring + " from 产品列表").Tables[0];
            return templist;
        }

        public static List<string> foldername_list()
        {
            List<string> templist = fullname_list("concat(名称,'_',图号)");
            return templist;
        }


        public static List<string> name_list(string para="001",string filternum="")
        {
            List<string> templist = new List<string>();

            switch(para)
            {
                case "process":
                    templist = DbHelperSQL.getlist("select replace(图号,'-001','process') from 产品列表 where 图号 like '%"+filternum+"%'");
                    break;

                case "null":
                    templist = DbHelperSQL.getlist("select replace(图号,'-001','') from 产品列表 where 图号 like '%"+filternum+"%'");
                    break;


                case "001":
                    templist = DbHelperSQL.getlist("select 图号 from 产品列表 where 图号 like '%"+filternum+"%'");
                    break;





            }
          
            return templist;
        }

       //查询某个产品的编号

       public static string queryno(string productname,string querystr)
        {
            string prodname = trimprod(productname);

            return DbHelperSQL.getlist("select " + querystr + " from 产品列表 where 图号 like '" + prodname + "%'").First();


        }

        #endregion
    

        //对于产品流水表的操作
        #region Production_op
        
       public static DataTable production_view()
       {
          return DbHelperSQL.Query("select 图号,站位号,产品名称,产品架次,开始日期,结束日期,移交日期,当前状态,备注,流水号 as id,状态说明 from 产品流水表 a left join 产品列表 b on a.产品名称=b.名称 ORDER BY 开始日期").Tables[0];
       }

        public static List<string> productionname_list(string para = "001")
        {
            List<string> templist = new List<string>();

            switch (para)
            {
                case "process":
                    templist = DbHelperSQL.getlist("select replace(图号,'-001','process') from 产品流水表 a left join 产品列表 b on a.产品名称=b.名称");
                    break;

                case "null":
                    templist = DbHelperSQL.getlist("select replace(图号,'-001','') from 产品流水表 a left join 产品列表 b on a.产品名称=b.名称");
                 
                    break;


                case "001":
                    templist = DbHelperSQL.getlist("select 图号 from 产品流水表 a left join 产品列表 b on a.产品名称=b.名称");
                   // templist = DbHelperSQL.getlist("select 图号 from 产品流水表");
                    break;

                   
                // startdate = false;



            }

            return templist;
        }



        #endregion




       //人员情况表的操作

       public static void people_table()
       {
          // DbHelperSQL.ExecuteSql("Create table 人员情况(单元 varchar(50) ,性质 varchar(50) ,类别 varchar(50) ,工资编号 varchar(50) ,姓名 varchar(50) PRIMARY KEY,用户名 varchar(50),密码 varchar(100),权限 int,手机 varchar(20),邮箱  varchar(50),IP  varchar(50),域名 varchar(50),WIN账户 varchar(50));");
            DbHelperSQL.ExecuteSql("CREATE TABLE PEOPLE(SECTION varchar(50) ,POSITION varchar(50),ID varchar(50),NAME varchar(50),USERNAME varchar(50),KEYCODE varchar(100),PRIVILEGE int,MOBILEPHONE varchar(20),EMAIL  varchar(50),NOTE text,primary key (ID,NAME));");
           // DbHelperSQL.ExecuteSql("ALTER TABLE PEOPLE ADD UNIQUE(ID);");


            MessageBox.Show("创建成功");

        }

      











       //对于RNC表的操作


       public static DataTable  RNC_view()
       {
           return DbHelperSQL.Query("select 外部拒收号,内部拒收号,拒收原因,纠正措施,当前状态,文件,关联产品,产品架次,发生日期,关闭日期,流水号,原因类型,责任人 from RNC总表 order by 发生日期").Tables[0];
       }



        //操作任务列表

        #region Task_op
       public static void update_taskprop(string index,string prop)
        {
            DbHelperSQL.ExecuteSql("update 任务管理 set 任务状态='"+prop+"' where 流水号='"+index+"'");
      //      MessageBox.Show("执行成功");

        }



        #endregion





        //对于试片列表的操作

        #region Coupon_op



        public static List<string> rfcouponno(string prodname)
        {
          string  productnametrim = prodname.Replace("-001", "Process");
            List<string> sqllist = new List<string>();
            //删除试片列表中存在的已被删除程序段的试片
            var segincoupons = DbHelperSQL.getlist("select 程序段编号 from 试片列表 where 产品图号='" + prodname+"';");
            var seginprocess= DbHelperSQL.getlist("select UUID from " + productnametrim);

          var toDel=  from aa in segincoupons
            where !seginprocess.Contains(aa)
            select aa;

          foreach (string pp in toDel)
            {
                sqllist.Add("delete from 试片列表 where 产品图号='" + prodname + "' and 程序段编号='" + pp+"';");
            }


            string sqlstrcp = "update 试片列表 aa inner join (SELECT @rownum := @rownum + 1 as rank,产品图号,蒙皮厚度,二层材料,二层厚度,totaltk from (SELECT  (二层厚度+蒙皮厚度) as totaltk,产品图号,蒙皮厚度,二层材料,二层厚度 FROM 试片列表 where 产品图号='" + prodname + "' group by 蒙皮厚度,二层材料,二层厚度 order by totaltk) bb,(SELECT @rownum := 0) r) kk on aa.二层厚度=kk.二层厚度 and aa.蒙皮厚度=kk.蒙皮厚度 and aa.二层材料=kk.二层材料 and aa.产品图号=kk.产品图号  set aa.编号=kk.rank";
            sqllist.Add(sqlstrcp);

            sqllist.Add("Update " + productnametrim + " pp left join (Select  程序段编号,GROUP_CONCAT(CONCAT('T',编号)) as 试片编号 from 试片列表 where 产品图号='" + prodname + "' group by 程序段编号) ll on pp.UUID=ll.程序段编号 set pp.试片Coupon_used=ll.试片编号");

            return sqllist;
        }












        //统计试片列表中铆钉需要用的数量

        public static  Dictionary<string,int> coupon_rivetqty(string productname)
        {

         productname=trimprod( productname)+"-001";
            Dictionary<string, int> rivetdic = new Dictionary<string, int>();
            Dictionary<string, int> hilitedic = new Dictionary<string, int>();
         // Dictionary<string, int> hilitedic = new Dictionary<string, int>();
            List<string> coupfst = DbHelperSQL.getlist("select 程序段编号 from 试片列表 where 产品图号='" + productname + "' and 程序段编号 not like '%B020600%'");
         foreach(string pp in coupfst)
         {
              //不是正常命名者不予添加试片数量

               if(pp.Split('_').Count()>1)
                {

             
             string fstname = pp.Split('_')[1];


             if (rivetdic.Keys.Contains(fstname))
                 {
                     rivetdic[fstname] = rivetdic[fstname] + 10;
                 }
             else
             {
                 rivetdic.Add(fstname, 15);

             }


                }


            }


         return rivetdic;




        }
     //统计试片列表中不同直径高锁需要用的数量
     public static Dictionary<string, int> coupon_hiliteqty(string productname)
     {

         productname = trimprod(productname) + "-001";
       //  Dictionary<string, int> rivettmpdic = new Dictionary<string, int>();
         Dictionary<string, int> hilitedic = new Dictionary<string, int>();

         
         // Dictionary<string, int> hilitedic = new Dictionary<string, int>();
         List<string> coupfst = DbHelperSQL.getlist("select 程序段编号 from 试片列表 where 产品图号='" + productname + "' and 程序段编号 like '%B020600%'");
         foreach (string pp in coupfst)
         {
             string fstname = pp.Split('_')[1];

           
                 if (hilitedic.Keys.Contains(fstname))
                 {
                     hilitedic[fstname] = hilitedic[fstname] + 5;
                 }
             else
                 {
                     hilitedic.Add(fstname, 10);
                     
                 }


             


         }

         return hilitedic;




     }

        #endregion






        //对于零件列表的操作
        #region Partlist_op

        public static void part_table()
        {
            //产品图号带-001
            DbHelperSQL.ExecuteSql("Create table if not exists 零件列表(编号 int,零件名称 varchar(50),零件图号 varchar(20),零件数量 int,备用数量 int,产品图号 varchar(50),有效架次 varchar(50),primary key (零件图号,产品图号));");
            MessageBox.Show("执行成功");
        }


     //从TVA表中导入并更新标准件 
       public static List<string> update_partlist()
        {

            List<string> prodlst = name_list();
            List<string> sqllist = new List<string>();
            List<string> buglist = new List<string>();


           foreach(string pp in prodlst)
           {
                //分析紧固件的备用数量


                DataTable fstlist=new DataTable();

                try
                {
                  fstlist = TVA_fstlist(pp, "INSTALLED BY");
                }
                catch
                {
                    if (!buglist.Contains(pp))
                    {
                        buglist.Add(pp);


                    }
                   //跳过这次循环
                    continue;
                }

               Dictionary<string,int> rivetdic=coupon_rivetqty(pp);

                Dictionary<string,int> hilitedic=coupon_hiliteqty(pp);

                //首先删除所有该表出现的紧固件

                sqllist.Add("delete from 零件列表 where 产品图号='" + pp + "' and (零件名称='HI-LITE' or 零件名称='RIVET')");

               for (int i=0;i<fstlist.Rows.Count;i++)
               {
                   string fstnum = fstlist.Rows[i][0].ToString();
                   string fstqty = fstlist.Rows[i][1].ToString();
                    int fstqtyint = System.Convert.ToInt16(fstqty);
                    string fstname ;
                   //备用耗损数量
                   int overqty=0;
                   if(fstnum.Contains("B020600"))
                   {
                       //按比例来分配
                 
                       //提取直径相同的高锁
                       string fstnametrunk=fstnum.Split('-')[0];
                       //统计高锁总数量
            
                int hiliteqty =fst_qty(pp,fstnametrunk);
                       fstname = "HI-LITE";
                       try
                       {
                            if (hilitedic.Keys.Contains(fstnametrunk))
                            {
                                overqty = (int)Math.Ceiling((double)fstqtyint / hiliteqty * hilitedic[fstnametrunk]);
                            }
                            else
                            {
                                overqty = 3;
                            }


                        
                       }
                       catch
                       {
                           //MessageBox.Show(pp + "TVA未更新或程序规划有问题");
                           if (!buglist.Contains(pp))
                           {
                               buglist.Add(pp);
                           }
                            //跳过这次循环
                           break;
                        }
                      
                      
                   }
                   else
                   {
                       fstname = "RIVET";
                       try
                       {
                            if (rivetdic.Keys.Contains(fstnum))
                            {
                                overqty = rivetdic[fstnum];
                            }
                          else
                            {
                                overqty = 5;
                            }



                        
                       }
                       catch
                       {
                          // MessageBox.Show(pp + "TVA未更新或程序规划有问题");
                           if (!buglist.Contains(pp))
                           {
                               buglist.Add(pp);
                           }
                            //跳过这次循环
                            break;
                        }

                   }

                   //这里对紧固件数量大于200的紧固件的紧固件进行补偿

                    if (fstqtyint > 200)
                    {
                        overqty = overqty + 10;
                    }

                   StringBuilder strSqlname = new StringBuilder();

                   strSqlname.Append("insert into 零件列表 (产品图号,零件图号,零件数量,零件名称,备用数量) values ");
                   strSqlname.Append(string.Format("('{0}','{1}',{2},'{3}',{4}) ON DUPLICATE KEY UPDATE 零件数量={2},备用数量={4}", pp, fstnum, fstqty, fstname, overqty));
                   sqllist.Add(strSqlname.ToString());
               }





           }

           DbHelperSQL.ExecuteSqlTran(sqllist);

           return buglist;
            
        }

       //返回全部紧固件数量
       public static DataTable spfsttable()
       {
           DataTable fsttemp = DbHelperSQL.Query("select 零件图号,sum(零件数量) as 数量,零件名称 from 零件列表 WHERE 零件名称 = 'RIVET' or 零件名称 = 'HI-LITE' group by 零件图号 order by 零件名称 desc,零件图号").Tables[0];
           return fsttemp;
       }
       //返回单个产品的紧固件数量
       public static DataTable spfsttable(string prodname)
       {
           prodname = trimprod(prodname);

           DataTable fsttemp = DbHelperSQL.Query("select 零件图号,零件数量 as 数量,零件名称 from 零件列表 WHERE 产品图号 like '" + prodname + "%' and (零件名称 = 'RIVET' or 零件名称 = 'HI-LITE') group by 零件图号 order by 零件名称 desc,零件图号").Tables[0];
           return fsttemp;
       }

        public static DataTable spfproctable(string prodname)
        {
            prodname = trimprod(prodname);
            try
            {
                DataTable fsttemp = DbHelperSQL.Query("select ProcessType as 加工类型,FastenerName as 紧固件,count(*) as 数量 from " + prodname + " group by 加工类型,紧固件 order by 加工类型,紧固件").Tables[0];
                return fsttemp;
            }
            catch
            {
                return new DataTable();
            }
           
        }

        //返回总共紧固件数量
        public static DataTable allqtytable(string prodname)
       {
           prodname = trimprod(prodname);

           DataTable fsttemp = DbHelperSQL.Query("select 零件图号,(sum(备用数量)+sum(零件数量)) as 数量,零件名称 from 零件列表 WHERE 产品图号 like '" + prodname + "%' and (零件名称 = 'RIVET' or 零件名称 = 'HI-LITE') group by 零件图号 order by 零件名称 desc,零件图号").Tables[0];
           return fsttemp;


       }
        public static DataTable allqtytable()
        {
           

            DataTable fsttemp = DbHelperSQL.Query("select 零件图号,(sum(备用数量)+sum(零件数量)) as 数量,零件名称 from 零件列表 WHERE  (零件名称 = 'RIVET' or 零件名称 = 'HI-LITE') group by 零件图号 order by 零件名称 desc,零件图号").Tables[0];
            return fsttemp;

        }

        //返回试片、吹钉的耗损数量
        public static DataTable overqtytable(string prodname)
       {
           prodname = trimprod(prodname);

           DataTable fsttemp = DbHelperSQL.Query("select 零件图号,sum(备用数量) as 数量,零件名称 from 零件列表 WHERE 产品图号 like '" + prodname + "%' and (零件名称 = 'RIVET' or 零件名称 = 'HI-LITE') group by 零件图号 order by 零件名称 desc,零件图号").Tables[0];
           return fsttemp;
      
       }
       public static DataTable overqtytable(List<string> prodlist)
       {
           DataTable fsttemp=new DataTable();
           foreach(string pp in prodlist)
           {
             string  prodname = trimprod(pp);

              fsttemp.Merge(DbHelperSQL.Query("select 零件图号,sum(备用数量) as 数量,零件名称 from 零件列表 WHERE 产品图号 like '" + prodname + "%' and (零件名称 = 'RIVET' or 零件名称 = 'HI-LITE') group by 零件图号 order by 零件名称 desc,零件图号").Tables[0]);
        
           }


           var query = from t in fsttemp.AsEnumerable()
                       group t by new { t1 = t.Field<string>("零件图号"),t2 = t.Field<string>("零件名称") } into m
                       select new
                       {
                           零件图号 = m.Key.t1,

                           数量 = m.Sum(n => n.Field<decimal>("数量")),
                           零件名称 = m.Key.t2
                       };
           DataTable newtemp = fsttemp.Clone();
           if (query.ToList().Count > 0)
           {
               query.ToList().ForEach(q =>
               {
                   newtemp.Rows.Add(new object[] {q.零件图号, q.数量, q.零件名称 });
                 //  Console.WriteLine(q.name + "," + q.sex + "," + q.score);
               });
           }
          
           

           return newtemp;

       }
        public static DataTable spfQtyStaticTable(Func<string,DataTable> querySpf, List<string> prodlist)
        {
            DataTable fsttemp = new DataTable();
            foreach (string pp in prodlist)
            {
                string prodname = trimprod(pp);

                fsttemp.Merge(querySpf(prodname));

            }


            var query = from t in fsttemp.AsEnumerable()
                        group t by new { t1 = t.Field<string>("零件图号"), t2 = t.Field<string>("零件名称") } into m
                        select new
                        {
                            零件图号 = m.Key.t1,

                            数量 = m.Sum(n =>System.Convert.ToInt32( n["数量"])),
                            零件名称 = m.Key.t2
                        };
            DataTable newtemp = fsttemp.Clone();
            if (query.ToList().Count > 0)
            {
                query.ToList().ForEach(q =>
                {
                    newtemp.Rows.Add(new object[] { q.零件图号, q.数量, q.零件名称 });
                    //  Console.WriteLine(q.name + "," + q.sex + "," + q.score);
                });
            }



            return newtemp;

        }
        public static DataTable overqtytable()
       {
          // prodname = trimprod(prodname);

           DataTable fsttemp = DbHelperSQL.Query("select 零件图号,sum(备用数量) as 数量,零件名称 from 零件列表 WHERE 零件名称 = 'RIVET' or 零件名称 = 'HI-LITE' group by 零件图号 order by 零件名称 desc,零件图号").Tables[0];
           return fsttemp;

       }

       //返回零件列表
       public static DataTable parttable(string prodname)
       {
           prodname = trimprod(prodname);

           DataTable fsttemp = DbHelperSQL.Query("select 零件图号,零件数量,零件名称 from 零件列表 WHERE 产品图号 like '" + prodname + "%' and (零件名称 <> 'RIVET' and 零件名称 <> 'HI-LITE')  order by 零件图号").Tables[0];
           return fsttemp;
       }

       public static DataTable parttable()
       {
         //  prodname = trimprod(prodname);

           DataTable fsttemp = DbHelperSQL.Query("select * from 零件列表 WHERE 零件名称 <> 'RIVET' and 零件名称 <> 'HI-LITE'  order by 零件图号").Tables[0];
           return fsttemp;
       }

        #endregion



       //对于零件库存表格的操作


       public static void partstore_table()
       {
           //产品图号带-001
           DbHelperSQL.ExecuteSql("Create table if not exists store_state(零件号 varchar(50),名称 varchar(50),单机数 int,结存 int,工位号 varchar(50),最后架次 varchar(50),primary key (零件号,工位号));");
           MessageBox.Show("执行成功");
       }

                
      public static DataTable loadfromExcel()
        {
          
            var fileDialog = new OpenFileDialog();

            fileDialog.InitialDirectory = "D://";

            fileDialog.Filter = "xls files (*.xls)|*.xls;*.xlsx|All files (*.*)|*.*";

            fileDialog.FilterIndex = 1;

            fileDialog.RestoreDirectory = true;

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {


                return OFFICE_Method.excelMethod.LoadDataFromExcel(fileDialog.FileName);
            }

            return null;
            }

    public static void updateTable(Func<DataRow,string> sqlGen)
        {

            var testExcel = new System.Data.DataTable();
            var fileDialog = new OpenFileDialog();

            fileDialog.InitialDirectory = "D://";

            fileDialog.Filter = "xls files (*.xls)|*.xls;*.xlsx|All files (*.*)|*.*";

            fileDialog.FilterIndex = 1;

            fileDialog.RestoreDirectory = true;

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {


                testExcel = OFFICE_Method.excelMethod.LoadDataFromExcel(fileDialog.FileName);




                var shuju = testExcel.Select();


                var creatName = new List<string>();

              


                foreach (DataRow p in shuju)
                {



                    if (p[0].ToString() != "")

                    {

                       


                        creatName.Add(sqlGen(p));
                    }


                }


                var k = DbHelperSQL.ExecuteSqlTran(creatName);

                MessageBox.Show(String.Format("执行成功,增加 '{0}'条记录", k));



            }




        }


        //对于紧固件列表的操作

        #region Fastener_table

        
        public static void ini_fasttable()

        {
            fast_table();
            Func<DataRow, string> sql = delegate (DataRow p)
              {
                  var strSqlname = new StringBuilder();
                  strSqlname.Append("REPLACE INTO 紧固件列表 VALUES(");



                  strSqlname.Append(String.Format("'{0}',{1},{2},'{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}')", p[0].ToString(), p[1].ToString(), p[2].ToString(), p[3].ToString(), p[4].ToString(), p[5].ToString(), p[6].ToString(), p[7].ToString(), p[8].ToString(), p[9].ToString(), p[10].ToString(), p[11].ToString(), p[12].ToString(), p[13].ToString(), p[14].ToString(), p[15].ToString(), p[16].ToString(), p[17].ToString(), p[18].ToString()));



                 return  strSqlname.ToString();
              };


            updateTable(sql);

        }


       public static void fast_table()
       {
           DbHelperSQL.ExecuteSql("Create table if not exists 紧固件列表(Fasteners varchar(100) PRIMARY KEY,Diameter double,Grip_max double,Tcode varchar(20),Cycle_Type varchar(20),Resync_Target varchar(20),Drill varchar(40),Upper_Anvil varchar(20),Tips varchar(30),Process_NO varchar(20),Countersink_depth varchar(20),Speed_of_drill varchar(20),Feed_speed varchar(20),Clamp_force varchar(20),Clamp_relief_force varchar(20),Upset_force varchar(20),Upset_position varchar(20),Seal_pres varchar(20),Seal_time varchar(20));");
  

       }
    
       public static List<string> allfast_list()
       {
           return DbHelperSQL.getlist("select Fasteners from 紧固件列表");
       }




              #endregion

       //数模表格

       #region Catia_model



       public static void catiaModel_table()
       {

           DbHelperSQL.ExecuteSql("Create table if not exists 产品数模(文件名 varchar(100) NOT NULL PRIMARY KEY,文件类型 varchar(50),历史 longtext,签名 text,产品名称 varchar(100),产品架次 varchar(100),修改日期 date,备注 text,文件地址 varchar(200));");

           MessageBox.Show("执行成功");
       }


       public static DataTable catiaModel_show()
       {
           return DbHelperSQL.Query("select * from 产品数模").Tables[0];


       }


       #endregion











        //创建表格




        #region creat_table


      //建立日志表

       public static void everyday_table()
       {
           DbHelperSQL.ExecuteSql("Create table if not exists everyDay(日期 date NOT NULL PRIMARY KEY,日期类型 varchar(100),事件 varchar(100),记录人 varchar(100),备注 text);");

           MessageBox.Show("执行成功");
       }

       public static void paperwork_table()
        {
            DbHelperSQL.ExecuteSql("Create table if not exists PaperWork(文件名 varchar(100) NOT NULL PRIMARY KEY,文件类型 varchar(100),文件编号 varchar(100),文件名称 varchar(100),版次 varchar(20),编制日期 date default null ,修改日期 date default null ,文件地址 varchar(100),文件夹 varchar(100),关联产品 varchar(100),关联架次 varchar(100),相关文件 varchar(100),文件格式 varchar(100),备注 text);");

            MessageBox.Show("执行成功");
        }

         public static void coupon_table()
        {
            DbHelperSQL.ExecuteSql("Create table if not exists 试片列表(编号 int,蒙皮厚度 int,二层材料 varchar(20),二层厚度 int,产品图号 varchar(50),程序段编号 varchar(100),primary key (蒙皮厚度,二层材料,二层厚度,产品图号,程序段编号));");
            MessageBox.Show("执行成功");
        }



       //
                public static void task_table()
         {
             //产品图号带-001
             DbHelperSQL.ExecuteSql("Create table if not exists 任务管理(流水号 int(5) NOT NULL AUTO_INCREMENT PRIMARY KEY,责任人 varchar(50) ,生成时间 timestamp NOT NULL default CURRENT_TIMESTAMP,任务名称 varchar(100),任务类型 varchar(100),任务说明 longtext,关联产品 varchar(100),节点日期 date default null ,完成日期 date default null ,任务状态 varchar(100),绩效奖励 double,额外奖励 double);");

             MessageBox.Show("执行成功");
         }

             public static void taskmodel_table()
         {
             DbHelperSQL.ExecuteSql("Create table if not exists 任务模板(责任人 varchar(50) ,任务名称 varchar(100) NOT NULL PRIMARY KEY,任务类型 varchar(100),任务说明 longtext,绩效奖励 double);");

         }
    
       //推送信息表

             public static void pushinfo_table()
             {
                 DbHelperSQL.ExecuteSql("Create table if not exists 推送信息(流水号 int(5) NOT NULL AUTO_INCREMENT PRIMARY KEY,责任人 varchar(50),信息状态 varchar(50),信息 longtext,生成时间 timestamp NOT NULL default CURRENT_TIMESTAMP);");
                 MessageBox.Show("执行成功");
             }
                    public static void rnc_table()
             {

                 DbHelperSQL.ExecuteSql("Create table if not exists RNC总表(流水号 int(5) NOT NULL AUTO_INCREMENT PRIMARY KEY,外部拒收号 varchar(100),内部拒收号 varchar(100),拒收原因 longtext,纠正措施 longtext,关联产品 varchar(100),产品架次 varchar(100),发生日期 date default null ,关闭日期 date default null ,当前状态 varchar(100),文件 varchar(100),责任人 varchar(20));");

                 MessageBox.Show("执行成功");
             }








                           public static void production_table()
             {

                 DbHelperSQL.ExecuteSql("Create table if not exists 产品流水表(流水号 int(5) NOT NULL AUTO_INCREMENT PRIMARY KEY,产品名称 varchar(100),产品架次 varchar(100),开始日期 date,结束日期  date default null,移交日期 date default null ,备注 text,当前状态 varchar(50),状态说明 text;");

                 MessageBox.Show("执行成功");
             }



        #endregion

                           //log表


                           #region log

                           public static void log_table()
                           {

                               DbHelperSQL.ExecuteSql("Create table if not exists Log(seq int(5) NOT NULL AUTO_INCREMENT PRIMARY KEY,clientID varchar(100),behavior text,occurtime timestamp NOT NULL default CURRENT_TIMESTAMP,note text);");

                               MessageBox.Show("执行成功");


                           }

                           public static void log_insert(string id,string info)
                           {

                               DbHelperSQL.ExecuteSql("Insert into Log(clientID,behavior) values('"+id+"','"+info+"');");

                           // MessageBox.Show("执行成功");


                           }





                           #endregion




                           //Material表

       #region material

                           public static void material_table()
                           {

                               DbHelperSQL.ExecuteSql("Create table if not exists Material(流水号 int(5) NOT NULL AUTO_INCREMENT PRIMARY KEY,名称 varchar(100),数量 int,单位 varchar(100),规格 varchar(100),关联文件 varchar(100),备注 longtext,提出日期 date default null ,到货日期 date default null ,当前状态 varchar(100));");

                               MessageBox.Show("执行成功");
                           }
                           public static DataTable material_view()
                           {

                               return DbHelperSQL.Query("select SEQ as 流水号,名称,数量,单位,category as 类别,备注,当前状态,提出日期,到货日期 from Material").Tables[0];
                           }
        #endregion

        #region tools

        public static void tool_table()
        {

            DbHelperSQL.ExecuteSql("Create table if not exists Tools(SEQ int(5) NOT NULL AUTO_INCREMENT PRIMARY KEY,NAME varchar(100),MFR varchar(100),DWG varchar(100),CODE varchar(100),ORG_CODE varchar(100),QTY int,SHOP_QTY int,UNIT varchar(100),CATEGORY varchar(100),NOTE longtext,PREPARE_DATE date default null ,RECEIVE_DATE date default null ,STATE varchar(100));");

            MessageBox.Show("执行成功");
        }

        public static void updateToolsTable()
        {
            Func<DataRow, string> sql = delegate (DataRow p)
            {
              

                return "insert into Tools(NAME,QTY,UNIT,CATEGORY,NOTE,STATE,PREPARE_DATE,RECEIVE_DATE,SHOP_QTY,MFR,DWG,CODE,ORG_CODE) values('" +p["工具名称"].ToString() + "'," + p["需求"].ToString() + ",'" + p["单位"].ToString() + "','" + p["工具类别"].ToString() + "','" + p["备 注"].ToString() + "','','',''," + p["实发"].ToString() + ",'" + p["制造单位"].ToString() + "','" + p["规格型号"].ToString() + "','" + p["工具编号"].ToString() + "','" + p["出厂编号"].ToString() + "')" ;
            };


            updateTable(sql);
        }
        public static void updatePeopleTable()
        {
            Func<DataRow, string> sql = delegate (DataRow p)
            {


                return "REPLACE INTO People(SECTION,POSITION,ID,NAME,USERNAME,KEYCODE,PRIVILEGE,MOBILEPHONE,EMAIL,NOTE) values('" + p["SECTION"].ToString() + "','" + p["POSITION"].ToString() + "','" + p["ID"].ToString() + "','" + p["NAME"].ToString() + "','" + p["USERNAME"].ToString() + "','" + p["KEYCODE"].ToString() + "'," + p["PRIVILEGE"].ToString() + ",'" + p["MOBILEPHONE"].ToString() + "','" + p["EMAIL"].ToString() + "','" + p["NOTE"].ToString() + "')";
            };


            updateTable(sql);
        }



        public static DataTable tool_view()
        {

            return DbHelperSQL.Query("select SEQ as 流水号,NAME as 名称,MFR as 供应商,DWG as 图号,CODE as 工具编号,ORG_CODE as 出厂编号,QTY AS 数量,SHOP_QTY AS 现场数量,UNIT AS 单位,CATEGORY as 类别,NOTE AS 备注,STATE AS 当前状态,PREPARE_DATE AS 提出日期,RECEIVE_DATE AS 到货日期 from Tools").Tables[0];
        }


        #endregion

        #region trouble_shoot

        public static void trouble_table()
        {

            DbHelperSQL.ExecuteSql("Create table if not exists Troubles(SEQ int(5) NOT NULL AUTO_INCREMENT PRIMARY KEY,NAME varchar(100),CATEGORY varchar(100),NOTE longtext,OCCUR_DATE date default null ,SOLVE_DATE date default null ,STATE varchar(100));");

            MessageBox.Show("执行成功");
        }

        public static DataTable trouble_view()
        {

            return DbHelperSQL.Query("select SEQ as 流水号,NAME as 名称,CATEGORY as 类别,NOTE AS 备注,STATE AS 当前状态,OCCUR_DATE AS 发生日期,SOLVE_DATE AS 解决日期 from Troubles").Tables[0];
        }


        #endregion



        #region MBOM

        public static void EBOM_table()
        {
            //产品图号带-001
            DbHelperSQL.ExecuteSql("Create table if not exists MBOM(PARTNUMBER varchar(50),NHA varchar(50),NAME int,MATERIAL int,PARTSTATE varchar(50),AOI varchar(50),CLASS varchar(50),LOCATETYPE varchar(50),SUPPLY varchar(50),DIVISION  varchar(50),NOTE text,primary key (PARTNUMBER,AOI));");
            MessageBox.Show("执行成功");
        }

        public static void UpdateMBOMtable()

        {

            Func<DataRow, string> sql = delegate (DataRow p)
            {
                var strSqlname = new StringBuilder();
                strSqlname.Append("INSERT INTO MBOM VALUES(");



                strSqlname.Append(String.Format("'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}') ON DUPLICATE KEY UPDATE DIVISION='{9}',CLASS='{6}'", p[0].ToString(), p[1].ToString(), p[2].ToString().Replace("\'",""), p[3].ToString(), p[4].ToString(), p[5].ToString(), p[6].ToString(), p[7].ToString(), p[8].ToString(), p[9].ToString(), p[10].ToString()));



                return strSqlname.ToString();
            };


            updateTable(sql);

        }

        #endregion

    }
}
