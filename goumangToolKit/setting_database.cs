using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GoumangToolKit
{
    public partial class setting_database : Form
    {
        public setting_database()
        {
            InitializeComponent();
        }
     public   string  addr
        {
            get
            {
                return textBox3.Text;

            }
         set
         {
             textBox3.Text = value;
             
         }
            
        }

     public string user
     {
         get
         {
             return textBox1.Text;

         }
         set
         {
             textBox1.Text = value;

         }

     }
     public string pwd
     {
         get
         {
             return textBox2.Text;

         }
         set
         {
             textBox2.Text = value;

         }

     }
        private void button1_Click(object sender, EventArgs e)
        {
            string datastr = "Database='autorivet';Data Source='" + addr + "';User Id='" + user + "';Password='" + pwd + "';CharSet=utf8;Allow User Variables=True";
            Properties.Settings.Default.datastr = datastr;
            PubConstant.ConnectionString = datastr;
            Properties.Settings.Default.Save();
            this.Close();
        }

        private void setting_database_Load(object sender, EventArgs e)
        {
            string constr = PubConstant.ConnectionString;
            addr = constr.Split(';')[1].Split('\'')[1].Replace("'", "");
         user = constr.Split(';')[2].Split('\'')[1].Replace("'", "");
           pwd = constr.Split(';')[3].Split('\'')[1].Replace("'", "");
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.SelectedIndex!=-1)
            {
                addr = comboBox1.Text.ToString();
            }
        }
    }
}
