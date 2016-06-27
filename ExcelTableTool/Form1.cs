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
            int filecount;

            int currentCount;
            bool isSuccess;
            try
            {
                //获取excel表中的数据
                DataTable tableSheet = removeTable.GetExcelTable(textBox1.Text, true);
                textBox3.Text = "0/" + tableSheet.Rows.Count.ToString();
                //分组之后放入不同的dt中然后放到List中

                List<DataTable> dtList = removeTable.GetGroupSendMemberTodt(tableSheet, out currentCount);

                textBox3.Text = currentCount.ToString() + "/" + tableSheet.Rows.Count.ToString();
                //把分组之后的表都存储到新建的excel中

                isSuccess = removeTable.ExportToExcel(dtList, out filecount);
                if (isSuccess)
                {
                    string a = filecount.ToString();
                    MessageBox.Show(string.Format("拆表成功，共导出{0}个文件", filecount));
                }
                else
                {
                    MessageBox.Show(string.Format("拆表有失败的，共成功导出{0}个文件", filecount));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("请求失败" + ex.ToString());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog();
            o.ShowDialog();
            textBox1.Text = o.FileName;
        }
    }
}
