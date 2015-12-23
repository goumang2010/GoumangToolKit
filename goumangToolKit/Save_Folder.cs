using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace mysqlsolution
{
    public partial class Save_Folder : Form
    {
        public Save_Folder()
        {
            InitializeComponent();
        }
        string SavePath = "";
        private void browser_Click(object sender, EventArgs e)
        {
           //  string     SavePath="";
        FolderBrowserDialog  brow  = new FolderBrowserDialog();
        brow.SelectedPath = Properties.Settings.Default.info_path;
        if( brow.ShowDialog()==DialogResult.OK)
        {
            SavePath = brow.SelectedPath;
           
            textBox1.Text = SavePath;

        }



        }

        private void Save_Folder_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.info_path;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.info_path= SavePath;
            Properties.Settings.Default.Save();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Reset();
        }
    }
}
