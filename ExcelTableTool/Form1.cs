using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelTableApi.Api.impl;
using ExcelTableApi.Api.service;

namespace ExcelTableTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private IRemoveTableBLL removeTable = new RemoveTableBLL();
           

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("请导入源数据");
                return;
            }
            //获取excel表中的数据
            DataTable tableSheet = removeTable.GetExcelTable(textBox1.Text, true);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog();
            o.ShowDialog();
            textBox1.Text = o.FileName;
        }
    }
}
