using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorkbook4
{
    public partial class Sheet4
    {
        private void Sheet4_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet4_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.button2.Click += new System.EventHandler(this.button2_Click);
            this.Startup += new System.EventHandler(this.Sheet4_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet4_Shutdown);

        }


        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "请选择文件";
            openFileDialog1.Filter = "所有文件(*.*)|*.*|(*.xlsm)|*.xlsm|(*.xlsx)|*.xlsx|(*.xls)|*.xls";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Globals.Sheet4.get_Range("B8").Value2 = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "请选择文件";
            openFileDialog1.Filter = "所有文件(*.*)|*.*|(*.xlsm)|*.xlsm|(*.xlsx)|*.xlsx|(*.xls)|*.xls";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Globals.Sheet4.get_Range("B10").Value2 = openFileDialog1.FileName;
            }
        }
    }
}
