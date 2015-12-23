using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using mysqlsolution;

namespace mysqlsolution
{
    public partial class productList : UserControl
    {
        public productList()
        {
            InitializeComponent();
        }

        private void productList_Load(object sender, EventArgs e)
        {
            listBox1.DataSource = AutorivetDB.foldername_list();
            listBox1.SelectedIndex = -1;

        }


    }
}
