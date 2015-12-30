using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using FileManagerNew;
using System.Security.Cryptography;

namespace mysqlsolution
{
    public partial class program_input : Form
    {


        public program_input()
        {
            InitializeComponent();
            DataTable duizhao = DbHelperSQL.Query("select Tcode, Fasteners,Resync_Target from 紧固件列表").Tables[0];
            int count = duizhao.Rows.Count;
            for (int i = 0; i < count; i++)
            {
                string Tcode = duizhao.Rows[i][0].ToString().Replace("T", "");
                fstenerT.Add(System.Convert.ToInt16(Tcode), duizhao.Rows[i][1].ToString());
                fstenerTR.Add(System.Convert.ToInt16(Tcode), duizhao.Rows[i][2].ToString());
            }
        }
        public List<string> program_part = new List<string>();
        string productnametrim;
        string prodchnname;
        Dictionary<string, int> wronglist = new Dictionary<string, int>();
        Dictionary<string, int> allfastlist = new Dictionary<string, int>();
        int  tvaqty = 0;
        bool ifoldprogram = true;

        Dictionary<string, int> alldrilllist = new Dictionary<string, int>();
        Dictionary<int, string> fstenerT = new Dictionary<int, string>();
        //T代码和校准代码对照
        Dictionary<int, string> fstenerTR = new Dictionary<int, string>();
        Dictionary<string, int> fastlist = new Dictionary<string, int>();
        Dictionary<string, int> drilllist = new Dictionary<string, int>();
        int tvadrillqty = 0;
      //  List<string> abc = new List<string>();
        int fstqty = 0;

        private void program_input_Load(object sender, EventArgs e)
        {



        }
        public string inputValue
        {
           

            set
            {

                if (value=="")
                {
                    menuStrip1.Visible = false;
                    button5.Visible = false;
                    comboBox1.Visible = false;
                    checkBox3.Visible = false;
                    label6.Visible = false;
                    label7.Visible = false;
                }
                else
                {
                    menuStrip1.Visible = true;
                  button5.Visible = true;
                    comboBox1.Visible = true;
                    checkBox3.Visible = true;

                   
                }
                loadtable(value);
               
            }
        }

        private void loadtable(string productname)
        {
            if(productname!="")
            {
                string prodnametrim = productname.Replace("process", "");
                this.Text = autorivet_op.queryno(prodnametrim, "程序编号");

                DataTable fsttemp = autorivet_op.TVA_fstlist(prodnametrim, "INSTALLED BY");


                DataTable fstdrilltemp = autorivet_op.TVA_fstlist(prodnametrim, "DRILL ONLY BY");
                
              //  DbHelperSQL.Query("select FastenerName,count(*) as qty from " + productname.Replace("process", "") + " WHERE ProcessType like '%DRILL ONLY BY%' group by FastenerName order by FastenerName").Tables[0];
            string fstdisplay = "TVA紧固件安装数量:\r\n";
            foreach(DataRow pp in fsttemp.Rows)
            {
                int tempqty=System.Convert.ToInt16(pp[1].ToString());
                fstdisplay = fstdisplay +pp[0]+ "  ： " + pp[1] + "\r\n";
                allfastlist.Add(pp[0].ToString(),tempqty);
                tvaqty = tvaqty + tempqty;
               // tvaqty=
            }
            fstdisplay = fstdisplay + "钻孔的(drill):\r\n";

            foreach (DataRow pp in fstdrilltemp.Rows)
            {
                int tempqty = System.Convert.ToInt16(pp[1].ToString());
                fstdisplay = fstdisplay + pp[0] + "  ： " + pp[1] + "\r\n";
                alldrilllist.Add(pp[0].ToString(), tempqty);

                tvadrillqty = tvadrillqty + tempqty;
                // tvaqty=
            }



            label7.Text = fstdisplay;

            try
            {

                productnametrim = productname;
                  prodchnname = autorivet_op.queryno(productnametrim, "名称");

                    MySqlConnection MySqlConn = new MySqlConnection(PubConstant.ConnectionString);
                MySqlConn.Open();
                String sql = "SELECT uuid FROM  " + productnametrim +" order by ID";
                program_part = DbHelperSQL.getlist(sql);


                comboBox1.DataSource = program_part;


                //   this.dataGridView1.DataSource = dt;

            }
            catch
            {
                MessageBox.Show("当前数据库不可用，请更换数据库");

            }

            }

            else
            {
                DataTable fsttemp = DbHelperSQL.Query("select Fasteners from 紧固件列表").Tables[0];
                foreach (DataRow pp in fsttemp.Rows)
                {
                    allfastlist.Add(pp[0].ToString(), 0);
                    alldrilllist.Add(pp[0].ToString(), 0);


                }

            }


        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ifoldprogram = true;

            if(comboBox1.SelectedIndex!=-1)
            {

         
            string pppart = comboBox1.SelectedValue.ToString();
                
           
             
                List<string> tem= DbHelperSQL.getlist("select 程序Program from " + productnametrim + " where UUID='" + pppart + "'").First().Split(new Char[2] { '\r', '\n' }, System.StringSplitOptions.RemoveEmptyEntries).ToList() ;
                listBox1.DataSource = tem;
                checkdupi(tem,false);
            }
            }
        private static bool installpoint(String s)
        {
            if (s.Contains("X") && s.Contains("Y") && s.Contains("Z"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool findtcode(String s)
        {
            if (s.Contains("M56"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool findgeoset(String s)
        {
            if (s.ToUpper().Contains("START GEOSET"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool findop(String s)
        {
            if (s.ToUpper().Contains("START OPERATION"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool installpointmm(String s)
        {
            if (s.Contains("M60") || s.Contains("M62"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool drillpointmm(String s)
        {
            if (s.Contains("M61") || s.Contains("M63"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private string creatpointinfo(string[] instcoodrow,int pp,List<string> abc)
        {
            string instcood = instcoodrow[0].Trim() + "," + instcoodrow[1].Replace(")", "").Replace("_", ",").Trim();
            //get the geoset
            int geosetindex = abc.FindLastIndex(pp - 1, findgeoset);
            string geosetstr = abc.ElementAt(geosetindex).Split(':')[1];
            geosetstr = geosetstr.Remove(geosetstr.Length - 1);
            geosetstr = geosetstr.Trim();

            //get the operation
            int opindex = abc.FindLastIndex(pp - 1, findop);
            string opstr = abc.ElementAt(opindex).Split(':')[1];
            opstr = opstr.Remove(opstr.Length - 1);
            opstr = opstr.Trim();

            return  instcood + "," + opstr + "," + geosetstr + "," + pp.ToString();
        
        }

        private string checkdupi( List<string> abc,bool output=true)
        {
           fastlist = new Dictionary<string, int>();

            int mm = abc.FindIndex(0, installpointmm); ;
                int pp = 0;
                Dictionary<string,int> installlist = new Dictionary<string,int> ();
                List<string> installlistfull = new List<string>();
                int uid = 0;

                while(mm!=-1)
                {
                   
                 //   行号
                   pp = abc.FindLastIndex(mm-1, installpoint);
                    string[] instcoodrow=abc.ElementAt(pp).Split('(');
                    
                    string instcoodsimple = instcoodrow[0].Trim();
                    //包括行号信息的ID
                    string instcood = creatpointinfo(instcoodrow,pp,abc);
                    //添加带行号信息的
                    installlistfull.Add(instcood);
                    string fastname = instcood.Split(',')[1];
                    string Tcode;
                    try
                    {
                       Tcode = abc[abc.FindLastIndex(mm - 1, findtcode)].Replace("M56T", "");
                    }
                    catch
                    {
                        MessageBox.Show("这不是完整的程序，没有包含选钉代码");

                        break;
                    }
                  
                    string fstnameT = fstenerT[System.Convert.ToInt16(Tcode)];
                    if (fstnameT != fastname)
                    {
                        string shuchu = "ProcessFeature名称错误," + instcood;
                       if( !wronglist.Keys.Contains(shuchu))
                       {
                           listBox2.Items.Add(shuchu);
                           wronglist.Add(shuchu, pp);
                       }


                    }
                    fastname = fstnameT;
               
                    if(fastlist.Keys.Contains(fastname))
                    {
                        fastlist[fastname]=fastlist[fastname]+1;

                    }
                    else
                    {
                        fastlist.Add(fastname,1);

                    }
                   // fastlist.Add(instcood[1],)
                   string  instcoodsimple2 = Regex.Replace(instcoodsimple, @"\.[0-9]*", "", RegexOptions.None);

                    if (installlist.Keys.Contains(instcoodsimple2) && (!listBox2.Items.Contains(instcood)))
                    {
                        instcood = uid + "," + instcood;
                        

                      
                        //找到上个重复值
                        string pppoint= installlistfull.Find(
                                          delegate(string bk)
              {

                  return bk.Contains(instcoodsimple);
              }
              );
                        int ppp = System.Convert.ToInt16((pppoint.Split(',')[5]));
                        pppoint = uid +","+ pppoint;

                        //如果还是找到同一个，则怀疑只是两点相近
                        if(instcood==pppoint)
                        {
                            string lastcoord = installlistfull[installlist[instcoodsimple2]].ToString();
                            //找到A值
                            Regex regex = new Regex(@"A-?[0-9]*\.[0-9]*");//我们要在目标字符串中找到"OK"
                            Match m1 = regex.Match(instcood);
                            Match m2 = regex.Match(lastcoord);
                            double mm1=System.Convert.ToDouble( m1.Value.Remove(0,1));
                            double mm2 = System.Convert.ToDouble(m2.Value.Remove(0, 1));
                            if(Math.Abs(mm1-mm2)<0.4)
                            {
                                pppoint = uid + "," + installlistfull[installlist[instcoodsimple2]];
                                wronglist.Add("坐标相近" + pppoint, ppp);
                                wronglist.Add(instcood, pp);
                                listBox2.Items.Add(pppoint);
                            }
                            
                          //  else
                           // {
                                //只是A相近不算疑似重复点
                              //  installlist.Add(instcoodsimple2, installlistfull.Count - 1);
                          //  }

                

                          
                               
                              // pppoint.Replace("," + ppp.ToString(), ","+installlist[instcoodsimple2].ToString());
                        }
                        else
                        {
                            wronglist.Add(pppoint, ppp);
                            wronglist.Add(instcood, pp);
                            listBox2.Items.Add(pppoint);
                        }

                       
                        

                        uid = uid + 1;



                    }


                    else
                    {
                        installlist.Add(instcoodsimple2, installlistfull.Count-1);

                    }
                    mm = abc.FindIndex(mm + 1, installpointmm);
                    
                }
                fstqty = installlistfull.Count();

            //Check drill points 2015.3.6

             drilllist = new Dictionary<string, int>();

                int mmd = abc.FindIndex(0, drillpointmm); ;
              //  int ppd = 0;
                List<string> drllist = new List<string>();
                List<string> drilllistfull = new List<string>();
                int uidd = 0;

                while (mmd != -1)
                {

                    //行号
                    pp = abc.FindLastIndex(mmd - 1, installpoint);
                    string[] instcoodrow = abc.ElementAt(pp).Split('(');

                    string instcoodsimple = instcoodrow[0].Trim();
                    string instcood = creatpointinfo(instcoodrow, pp, abc);

                    drilllistfull.Add(instcood);
                    installlistfull.Add(instcood);
                    string fastname = instcood.Split(',')[1];
                   // findtcode

                    string Tcode = abc[abc.FindLastIndex(mmd - 1, findtcode)].Replace("M56T", "");
                  string fstnameT=  fstenerT[System.Convert.ToInt16(Tcode)];
                   
                   



                    if(fstnameT!=fastname)
                    {
                        string shuchu = "ProcessFeature名称错误," + instcood;
                        if (!wronglist.Keys.Contains(shuchu))
                        {
                            listBox2.Items.Add(shuchu);
                            wronglist.Add(shuchu, pp);
                        }

                    }
                    fastname=fstnameT;
                    if (drilllist.Keys.Contains(fastname))
                    {
                        drilllist[fastname] = drilllist[fastname] + 1;

                    }
                    else
                    {
                        drilllist.Add(fastname, 1);

                    }
                    // fastlist.Add(instcood[1],)
                    if (installlist.Keys.Contains(instcoodsimple) && (!listBox2.Items.Contains(instcood)))
                    {
                           
                
                        instcood = uidd + "," + instcood;
                        wronglist.Add(instcood, pp);

                    
                            listBox2.Items.Add(instcood);
                       

                        string pppoint = installlistfull.Find(
                                          delegate(string bk)
                                          {

                                              return bk.Contains(instcoodsimple);
                                          }
              );
                        int ppp = System.Convert.ToInt16((pppoint.Split(',')[5]));
                        pppoint = uidd + "," + pppoint;

                        //钻孔未考虑相近点
                       
                        wronglist.Add(pppoint, ppp);
                        listBox2.Items.Add(pppoint);

                        uidd = uidd + 1;



                    }


                    else
                    {
                        installlist.Add(instcoodsimple, installlistfull.Count-1);

                    }
                    mmd = abc.FindIndex(mmd + 1, drillpointmm);

                }
            
            //drill analysis end












                if (checkBox2.Checked == true && output)
                {
                    //string[] titlefull = new string[6];

                    //titlefull[0] = "坐标Coord";
                    //titlefull[1] = "紧固件Fastener";
                    //titlefull[2] = "PF index";
                    //titlefull[3] = "Operation name";
                    //titlefull[4] = "Geoset name";
                    //titlefull[5] = "程序行号Program row";


                DataTable dt = new DataTable();
         
                dt.Columns.Add("坐标Coord", typeof(string));
                dt.Columns.Add("紧固件Fastener", typeof(string));
                dt.Columns.Add("PF index", typeof(string));
                dt.Columns.Add("Operation name", typeof(string));
                dt.Columns.Add("Geoset name", typeof(string));
                dt.Columns.Add("程序行号Program row", typeof(string));


                foreach (string item in installlistfull)
                {
                    dt.Rows.Add(item.Split(','));
                }
                if(dt.Rows.Count>0)
                {
                    OFFICE_Method.excelMethod.SaveDataTableToExcel(dt);
                }
              




            }

                if (output)
                {
                    //string[] title = new string[7];
                    //title[0] = "UID";
                    //title[1] = "坐标Coord";
                    //title[2] = "紧固件Fastener";
                    //title[3] = "PF index";
                    //title[4] = "Operation name";
                    //title[5] = "Geoset name";
                    //title[6] = "程序行号Program row";

                DataTable dt = new DataTable();
                dt.Columns.Add("UID", typeof(string));
                dt.Columns.Add("坐标Coord", typeof(string));
                dt.Columns.Add("紧固件Fastener", typeof(string));
                dt.Columns.Add("PF index", typeof(string));
                dt.Columns.Add("Operation name", typeof(string));
                dt.Columns.Add("Geoset name", typeof(string));
                dt.Columns.Add("程序行号Program row", typeof(string));

                foreach(string item in wronglist.Keys.ToList())
                {
                    dt.Rows.Add(item.Split(','));
                }
                if (dt.Rows.Count > 0)
                {
                    OFFICE_Method.excelMethod.SaveDataTableToExcel(dt);
                }
                

          // excelMethod.SaveListToExcel(title, wronglist.Keys.ToList());

                }


          
          int drlqty = drilllistfull.Count();

          string display = "NC代码紧固件数量：\r\n";
          foreach (var item in fastlist)
          {

              display = display + item.Key.ToString() + "： " + item.Value.ToString() + " TVA:"+allfastlist[item.Key.ToString()]+"\r\n";
              
          }
          display = display + "总安装数量：" + fstqty.ToString() + " TVA:" + tvaqty;
          display = display + "\r\nNC代码仅钻孔数量：\r\n";
          foreach (var item in drilllist)
          {
              try
                  {
                      display = display + item.Key.ToString() + "： " + item.Value.ToString() + " TVA:" + alldrilllist[item.Key.ToString()] + "\r\n";
                  }
              catch
              {
                  MessageBox.Show("请及时更新TVA");
              }
              

          }
          display = display + "总钻孔数量：" + drlqty.ToString() + " TVA:" + tvadrillqty;
          label6.Text = display;
            listBox1.DataSource = abc;
            return display;

                    }

        private List<string> creatlist(string text)
        {
            string[] rowprocess;
            List<string> abc = new List<string>();


            rowprocess = text.Split(new Char[2] { '\r', '\n' }, System.StringSplitOptions.RemoveEmptyEntries);

            foreach (string kk in rowprocess)
            {
                string tempstr;
                tempstr = kk.Trim();
                tempstr = Regex.Replace(tempstr, @"^N[0-9]*", "", RegexOptions.None);

                if (!tempstr.Contains("(MSG"))
                {
                    tempstr = tempstr.Replace(" ", "");
                }
                tempstr = tempstr.Trim();
               //2015.9.24 Remove M02 for each part
                if (tempstr == "M02")
                {
                    continue;
                }


                //修复换刀T代码bug
                if (tempstr.Contains("M56") && (!tempstr.Contains("T")))
                {
                    tempstr = tempstr.Replace("M56", "M56T");
                }
                
                //解决M34N/A bug(强制校准bug)
                if (!tempstr.Contains("M34N/A"))
                {
                    tempstr = tempstr.Replace("N/A", "");
                }
                else
                {
                    tempstr = "M34" + fstenerTR[System.Convert.ToInt16(abc.FindLast(findtcode).Replace("M56T", ""))];

                }

           


                if (tempstr != "")
                {
                    abc.Add(tempstr);
                }

            }
            return abc;
        }
        private void clearall()
        {
            wronglist.Clear();
           // listBox1.DataSource=null;
            
            listBox2.Items.Clear();
          //  abc.Clear();
            fstqty = 0;
            listBox2.DataSource = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            clearall();
            ifoldprogram = false;
            if (textBox1.Text!="")
            {
                List<string>  abc = creatlist(textBox1.Text);
                checkdupi(abc);
                listBox1.DataSource = abc;
                listtotext(abc);
                }
                
                 //abc.f

                // abc.FindLastIndex


            
    
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox1.DataSource = null;
            
            wronglist.Clear();
            listBox1.Items.Clear();
            listBox2.Items.Clear();
           // abc.Clear();
            //abc = new List<string>();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.ShowDialog();
            // MessageBox.Show(ofd.FileName);
            //异常检测开始
            try
            {
                FileStream fs = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read);//读取文件设定
                StreamReader m_streamReader = new StreamReader(fs, System.Text.Encoding.GetEncoding("GB2312"));//设定读写的编码
                //使用StreamReader类来读取文件
                m_streamReader.BaseStream.Seek(0, SeekOrigin.Begin);
                //  从数据流中读取每一行，直到文件的最后一行，并在rTB_Display.Text中显示出内容

                string strLine = m_streamReader.ReadToEnd();

                //关闭此StreamReader对象
                textBox1.Text = strLine;
                m_streamReader.Close();
            }
            catch
            {

            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex!=-1)
            {
                try
                {
                    int indexkey = wronglist[listBox2.SelectedItem.ToString()];
                    listBox1.SelectedIndex = indexkey;
                    listBox1.TopIndex = indexkey;
                }
                catch
                {

                }
                
            }
            
        }
        private string  listtotext(List<string> abc)
        {
            
            string newtext = "";

            foreach( string ddd in abc)
            {
                newtext = newtext + ddd + "\r\n";
            }
            textBox1.Text = newtext;
            return newtext;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            listBox1.TopIndex = listBox1.FindString(listBox1.SelectedItem.ToString());
            listBox1.SelectedIndex = listBox1.TopIndex;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            List<string> abc = (List<string>)listBox1.DataSource;
            abc.RemoveAt(listBox1.SelectedIndex);
            listBox1.DataSource = null;
            listBox1.DataSource = abc;
            listtotext(abc);
            listBox2.Items.Clear();
            //textBox1.Text=listBox1.Items.
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SaveFileDialog ofd = new SaveFileDialog();
            List<string> abc = (List<string>)listBox1.DataSource;
           if( ofd.ShowDialog() == DialogResult.OK)
           {

         
           
            StreamWriter sw = new StreamWriter(ofd.FileName, false);
            // sw.WriteLine(str);
            if (comboBox1.Visible==true&& checkBox3.Checked==true)
            {
                string firstrow = "(MSG,START PROGRAM PART " + (comboBox1.SelectedIndex + 1).ToString() + " :"+comboBox1.SelectedValue+")";
                
                sw.WriteLine(firstrow);
            }
            //string rivisetext = "";
            int indexNo = 2;
            if (checkBox1.Checked==true)
            {
                foreach (string ppp in abc)
                {
                    if (ppp.ElementAt(0) == 'X' || ppp.ElementAt(0) == 'M' || ppp.ElementAt(0) == 'G' || ppp.Contains("MSG"))
                    {
                        sw.WriteLine("N" + indexNo.ToString() + " " + ppp);
                        //rivisetext = rivisetext + "N" + indexNo.ToString() +ppp + "\r\n";
                        indexNo = indexNo + 2;
                    }
                    else
                    {
                        sw.WriteLine(ppp);
                    }
                }
            }
            else
            {

                foreach (string ppp in abc)
                {
                  
                        sw.WriteLine(ppp);
                    
                }

            }
            if (comboBox1.Visible == true && checkBox3.Checked == true)
            {
                string endrow = "(MSG,END PROGRAM PART " + (comboBox1.SelectedIndex + 1).ToString() + " :" + comboBox1.SelectedValue + ")";
                sw.WriteLine(endrow);
            }
            sw.Close();
            System.Diagnostics.Process pro = new System.Diagnostics.Process();
            pro.StartInfo.FileName = "notepad.exe";
            
            pro.StartInfo.Arguments = ofd.FileName;
            pro.Start();
            MessageBox.Show("生成完毕");

        }
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void checkAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            clearall();
            comboBox1.SelectedIndex = -1;
            listBox1.DataSource = null;
            List<string> abc = new List<string> ();
            int TVAqtyall = 0;
            int fstqtyall = 0;

            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                comboBox1.SelectedIndex = i;
                string value = comboBox1.SelectedValue.ToString();
                List<string> abc1 = (List<string>)listBox1.DataSource;
                abc.AddRange(abc1);

                // checkdupi(abc1,false);
              
        //   int TVAqty =(int) DbHelperSQL.Query("select 紧固件数量Fastener_Qty from " + productnametrim + " where uuid='" +value + "'").Tables[0].Rows[0][0];
                
        //  int TVAqtyBACKUP = (int)DbHelperSQL.Query("select 紧固件数量Fastener_Qty from " + productnametrim + "_backup where uuid='" + value + "'").Tables[0].Rows[0][0];

        //         if (fstqty != tvaqty)
       //     {
        //        listBox2.Items.Add("加工紧固件数量错误！程序ID：" + value + "；TVA数量：" + tvaqty + "；NC代码数量：" + fstqty);
         //   }
             


            }
            clearall();
            checkdupi(abc,true);
            comboBox1.SelectedIndex = -1;
            fstqtyall = fstqty;
         //   string questrtemp = "select sum(紧固件数量Fastener_Qty) from " + productnametrim + "_backup";
         //   TVAqtyall = System.Convert.ToInt32( DbHelperSQL.Query(questrtemp).Tables[0].Rows[0][0]);
        if (fstqtyall != TVAqtyall)
        {
            listBox2.Items.Add("加工紧固件总数量错误！TVA数量：" + tvaqty + "；NC代码数量：" + fstqtyall);

            MessageBox.Show("加工紧固件总数量错误！TVA数量：" + tvaqty + "；NC代码数量：" + fstqtyall);
        }
             
        }

       private void tiancpara()

       {
            string jingujian = "";

                string jingujianshuliang = "";

                Dictionary<string, string> outputlist = new Dictionary<string, string>();






                foreach (var item in fastlist)
                {
                    outputlist.Add(item.Key.ToString(), item.Value.ToString());

                   // jingujian = "["+jingujian + item.Key.ToString() + "： " + item.Value.ToString() + " TVA:" + allfastlist[item.Key.ToString()] + "\r\n";


                }
              
                foreach (var item in drilllist)
                {
                    string fstname=item.Key.ToString();
                    if (outputlist.Keys.Contains(fstname))
                    {
                        string oldqty= outputlist[fstname];
                        outputlist[fstname] = oldqty + "," + item.Value.ToString() ;
                    }
                    else
                    {
                        outputlist.Add(fstname, item.Value.ToString() );
                    }


                  //  display = display + item.Key.ToString() + "： " + item.Value.ToString() + " TVA:" + alldrilllist[item.Key.ToString()] + "\r\n";

                }
                List<List<string>> canshu = new List<List<string>>();

                foreach (var item in outputlist)
                {
                    string fstname = item.Key.ToString();
                   List<string> templist=(DbHelperSQL.getlistcol("select Process_NO,Countersink_depth,Speed_of_drill,Feed_speed,Clamp_force,Clamp_relief_force,Upset_force,Upset_position,Seal_pres,Seal_time from  紧固件列表 where Fasteners='" + fstname + "'"));

                   canshu.Add(templist);
                   if (outputlist.Count()==1)
                   {
                       jingujian = fstname;
                       jingujianshuliang = item.Value.ToString();
                   }
                   else
                   {
                       jingujian = jingujian + "[" + fstname + "]";

                       jingujianshuliang = jingujianshuliang + "[" + item.Value.ToString() + "]";
                       // jingujian = "["+jingujian + item.Key.ToString() + "： " + item.Value.ToString() + " TVA:" + allfastlist[item.Key.ToString()] + "\r\n";

                   }
                   // canshu.Add
                  
                }
                List<string> resultcanshu=new List<string> ();
                int fstcount=canshu.Count();
            if (fstcount==0)
            {
                return;
            }
                for (int jj=0;jj<10;jj++)
                    {
                        string tempstr="";
                    bool samepara=true;
                    string prestr = canshu[0][jj];
                      for(int dd=0;dd<fstcount;dd++)
                {


                      if(canshu[dd][jj]!=prestr) 
                      {
                          samepara=false;

                      }
                          prestr=canshu[dd][jj];

                  tempstr=tempstr+"[" + canshu[dd][jj]+"]";



                    }

                    if(samepara)
                    {
                        resultcanshu.Add(prestr);
                    }
                    else
                    {
                         resultcanshu.Add(tempstr);
                    }
                   


                }
                


                DbHelperSQL.ExecuteSql("Update " + productnametrim + " set 紧固件名称Fastener_Name='" + jingujian + "',紧固件数量Fastener_Qty='"+ jingujianshuliang+"',参数号Process_NO='"+resultcanshu[0]+"',锪窝深度Countersink_depth='"+resultcanshu[1]+"',钻头转速Speed_of_drill='"+resultcanshu[2]+"',给进速率Feed_speed='"+resultcanshu[3]+"',夹紧力Clamp_force='"+resultcanshu[4]+"',夹紧释放力Clamp_relief_force='"+resultcanshu[5]+"',墩铆力Upset_force='"+resultcanshu[6]+"',墩铆位置Upset_position='"+resultcanshu[7]+"',注胶压力Seal_pres='"+resultcanshu[8]+"',注胶时间Seal_time='"+resultcanshu[9]+"' where uuid='" + comboBox1.SelectedValue.ToString() + "'");
         
        
            
            
        
              //  MessageBox.Show("更新出错！");
              //  MessageBox.Show("更新出错！\r\n对于" + comboBox1.SelectedValue.ToString() + ":\r\nTVA中加工的紧固件数量：" + TVAqty + "\r\nNC代码中加工的紧固件数量：" + fstqty + "\r\n请考虑：" + "\r\n1.TVA是否为最新，更新TVA信息至数据库；\r\n2.检查NC代码是否有问题；\r\n3.直接修改Process表以忽略该问题。");
            

           
       }

        private void button5_Click(object sender, EventArgs e)
        {
            if (ifoldprogram)
            {
                MessageBox.Show("目前是旧程序，请重新点击Check");
            }
            else
            {
                clearall();
                List<string> abc = (List<string>)listBox1.DataSource;
                checkdupi(abc);


                // int TVAqty =(int) DbHelperSQL.Query("select 紧固件数量Fastener_Qty from " + productnametrim + " where uuid='" + comboBox1.SelectedValue.ToString() + "'").Tables[0].Rows[0][0];

                //    if(fstqty==TVAqty)
                //  {

                listtotext(abc);
                string temppro = textBox1.Text;
                tiancpara();

                //   }
                DbHelperSQL.ExecuteSql("Update " + productnametrim + " set 程序Program='" + temppro + "' where uuid='" + comboBox1.SelectedValue.ToString() + "'");

                MessageBox.Show("执行成功！");
            }


            
        }

        private void outputAllToolStripMenuItem_Click(object sender, EventArgs e)
        {

            //fetch the part quantity

          var partDic=  DbHelperSQL.getDic("select UUID,CONCAT(胶嘴Sealant_Tip,';',下铆头Lower_Anvil) from " + productnametrim );
            int partnum = 0;
            string lastpartname = "";
            try
            {
                List<string> tempabc = new List<string>();

                //string savefolder = Properties.Settings.Default.output_path + "\\" + productnametrim + "\\";
                string savefolder = Properties.Settings.Default.info_path +prodchnname+"_"+ productnametrim.Replace("process","-001") + "\\NC\\";

                if (!System.IO.Directory.Exists(savefolder))
                {
                    System.IO.Directory.CreateDirectory(savefolder);
                }
                //Do not look through sub dictionary
                List<FileInfo> rm = new List<FileInfo>();
                rm.WalkTree(savefolder, false);
              
                string newfolder = savefolder + "old";

                rm.moveto(newfolder);

                clearall();
                comboBox1.SelectedIndex = -1;
                //2015.9.4 change the name to the foemal program num.
                string filenameall = this.Text;


                string NCpath = savefolder + filenameall;


                StreamWriter swall = new StreamWriter(NCpath, false);



                //Backup and remove all the files in the folder





                //Look through all parts and output them
                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                
                    comboBox1.SelectedIndex = i;
                    tempabc = (List<string>)listBox1.DataSource;
                    if (tempabc.Count() != 0)
                    {

                        string uuid = comboBox1.SelectedValue.ToString();
                        int indexNo = 2;


                        string filename = "SEG_" + (i + 1).ToString() + "_" + uuid;
                        string filepath =savefolder + filename;
                        //if(System.IO.File.Exists(filepath))
                        //{
                        //    System.IO.File.Delete(filepath);
                        //}

                        StreamWriter sw = new StreamWriter(filepath, false);
                        Action<string> fillText = delegate (string text)
                        {
                            sw.WriteLine(text);
                            swall.WriteLine(text);
                        };



                        fillText("%");

                        fillText(productnametrim.Replace("process", "").Replace("C0", "O") + (comboBox1.SelectedIndex + 1).ToString());
                        if (partDic.Values.ElementAt(i) != lastpartname)
                        {
                            if (partnum != 0)
                            {
                                fillText("(Part" + partnum.ToString() + " END)");
                            }

                            partnum = partnum + 1;
                            fillText("(PART" + partnum.ToString() + " START)");
                            lastpartname = partDic.Values.ElementAt(i);


                        }



                        fillText("(MSG,START PROGRAM SEGMENT " + (comboBox1.SelectedIndex + 1).ToString() + " :" + uuid + ")");
                        fillText("(MSG,MAKE SURE THE SEALANT TIP AND LOWER ANVIL:" + partDic[uuid]+")");


                       // string value = comboBox1.SelectedValue.ToString();
                       tempabc.RemoveAt(0);
                        // tempabc.RemoveAt(0);
                        foreach (string ppp in tempabc)
                        {
                            if (ppp.ElementAt(0) == 'X' || ppp.ElementAt(0) == 'M' || ppp.ElementAt(0) == 'G' || ppp.Contains("MSG"))
                            {
                                string tmpstr = "N" + indexNo.ToString() + " " + ppp;
                                sw.WriteLine(tmpstr);
                                swall.WriteLine(tmpstr);
                                //rivisetext = rivisetext + "N" + indexNo.ToString() +ppp + "\r\n";
                                indexNo = indexNo + 2;
                            }
                            else
                            {
                                fillText(ppp);
                      
                            }
                        }

                     fillText( "(MSG,END PROGRAM SEGMENT " + (comboBox1.SelectedIndex + 1).ToString() + " :" + comboBox1.SelectedValue + ")");

                        //Add M02 by hand
                        sw.WriteLine("M02");
                      


                        sw.Close();

                    }
                }

                swall.WriteLine("M02");
                
                swall.Close();

               DateTime edittime= System.IO.Directory.GetLastWriteTime(savefolder);
                //DateTime nowtime=

               TimeSpan ts1 = new TimeSpan(edittime.Ticks);

               TimeSpan ts2 = new TimeSpan(DateTime.Now.Ticks);

               TimeSpan ts = ts1.Subtract(ts2).Duration();

               System.Diagnostics.Process.Start("explorer.exe", savefolder);
               if( ts.TotalSeconds>5)
               {
                   MessageBox.Show("输出失败");
               }
               else
               {
                    //Update the MD5 and the creation time in database

                    string fileMD5=  localMethod.GetMD5HashFromFile(NCpath);
                    string creationtime = edittime.ToString();
                    Dictionary<string, string> dic = new Dictionary<string, string>();
                    dic.Add("TIME", creationtime);
                    dic.Add("MD5", fileMD5);
                    string jsontext=   localMethod.ToJson(dic);
                  //  string prodchnname = autorivet_op.queryno(productnametrim, "名称");
                    DbHelperSQL.ExecuteSql("update 产品数模 set 备注='"+jsontext+"' where 产品名称='"+prodchnname+ "' and 文件类型='Process'");
                    MessageBox.Show("执行成功！已于"+ creationtime + "更新");




               }

               
            }
            catch(Exception ee)
            {
               
                MessageBox.Show("输出失败:"+ee.Message);
            }



            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            clearall();
        }

        private void databaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setting_database f1 = new setting_database();
            f1.Show();
           
        }

        private void savingFoldersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Save_Folder f1 = new Save_Folder();
            f1.Show();
            
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            List<string> abc = (List<string>)listBox1.DataSource;
            listtotext(abc);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {


        }

        private void fillParaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                comboBox1.SelectedIndex = i;
                string value = comboBox1.SelectedValue.ToString();
                List<string> abc1 = (List<string>)listBox1.DataSource;
                //abc.AddRange(abc1);

                checkdupi(abc1, false);

                tiancpara();



            }
            MessageBox.Show("执行成功！");
        }

        private void reTrimToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                comboBox1.SelectedIndex = i;
                string value = comboBox1.SelectedValue.ToString();
                List<string> abc1 = (List<string>)listBox1.DataSource;
                
            string newprogram=  listtotext (creatlist(listtotext(abc1)));

            DbHelperSQL.ExecuteSql("Update " + productnametrim + " set 程序Program='" + newprogram + "' where uuid='" + comboBox1.SelectedValue.ToString() + "'");


                //abc.AddRange(abc1);

               

              //  tiancpara();



            }
            MessageBox.Show("执行成功！");
        }
    }
}
