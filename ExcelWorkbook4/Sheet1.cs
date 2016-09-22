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
using Microsoft.Office.Interop.Excel;
using ClassApIBILL;
using System.Threading;
using System.Collections;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;


namespace ExcelWorkbook4
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
          
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
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
            this.button3.Click += new System.EventHandler(this.button3_Click);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion
        private void Sheet1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.D1&&e.Control)
            {
                Globals.Sheet1.get_Range("U").Value2 = 0;
            }
         }
        

        private void button1_Click(object sender, EventArgs e)
        {
            //Thread t = new Thread(new ThreadStart(NewMethod));
            //t.Start();
            NewMethod();
        }
        #region  
        public void NewMethod()
        {

            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            stopwatch.Start(); 
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
           string paixu=Globals.Sheet4.get_Range("B2").Value2;
            string dtbfs = Globals.Sheet4.get_Range("B4").Value2;
            int rowNumber = activeWorksheet.UsedRange.Rows.Count;
            int SNumber = Globals.Sheet3.UsedRange.Rows.Count;           
            string[] place = new string[SNumber];           
            //IndexerClass2 Indexer1 = new IndexerClass2();
            Application.StatusBar = "正在生成数组....";           
            for (int j = 0; j < SNumber; j++)
            {
                int k = j + 2;
                place[j] = Globals.Sheet3.get_Range("D" + k).Value2;
            }
            string[] name = new string[rowNumber];

            for (int k = 0; k < rowNumber; k++)
            {
                int l = k + 1;
                name[k] = activeWorksheet.get_Range("P" + l).Value2.ToString();
            }
            int time2 = 1;
            //object[,] mark = new string[1, rowNumber-2];

            if (dtbfs == "手动选择")
            {
                var selection = Globals.ThisWorkbook.Application.Selection;
                int a = selection.Rows.Count + 1;
                for (int i = 1; i < a; i++)
                {
                    if (selection.Cells.SpecialCells(XlCellType.xlCellTypeVisible, 12).Rows[i].Hidden == false)
                    {
                        int num = selection.Rows[i].Row;
                        if (Globals.Sheet1.get_Range("L" + num).Value2 == null)
                        {
                            continue;
                        }
                        string sendcity, receivercity;
                        string sendaddress = activeWorksheet.get_Range("N" + num).Value2;
                        string receiveraddress = activeWorksheet.get_Range("N" + num).Value2;
                        StringBuilder sb = new StringBuilder();
                        string province = activeWorksheet.get_Range("K" + num).Value2;
                        string city = activeWorksheet.get_Range("L" + num).Value2;
                        string district = activeWorksheet.get_Range("M" + num).Value2;
                        if (province == null || province == "")
                        {
                            sendcity = city + "," + city + "," + district;
                            receivercity = city + "," + city + "," + district;
                        }
                        else
                        {
                            sendcity = province + "," + city + "," + district;
                            receivercity = province + "," + city + "," + district;
                        }
                        GetDaTouBi(sendcity, sendaddress, receivercity, receiveraddress, sb);
                        activeWorksheet.get_Range("T" + num).Value2 = sb.ToString();
                    }
                }
            }
            else
            {
                for (int i = 2; i < rowNumber + 1; i++)
                {
                    string sendcity, receivercity;
                    string sendaddress = activeWorksheet.get_Range("N" + i).Value2;
                    string receiveraddress = activeWorksheet.get_Range("N" + i).Value2;
                    StringBuilder sb = new StringBuilder();
                    string province = activeWorksheet.get_Range("K" + i).Value2;
                    string city = activeWorksheet.get_Range("L" + i).Value2;
                    string district = activeWorksheet.get_Range("M" + i).Value2;
                    if (province == null || province == "")
                    {
                        sendcity = city + "," + city + "," + district;
                        receivercity = city + "," + city + "," + district;
                    }
                    else
                    {
                        sendcity = province + "," + city + "," + district;
                        receivercity = province + "," + city + "," + district;
                    }
                    string searchplace = province + city + district;                  
                    searchplace = Regex.Replace(searchplace, @"\s", "");                  
                    if (!((IList)place).Contains(searchplace))
                    {
                        GetDaTouBi(sendcity, sendaddress, receivercity, receiveraddress, sb);
                        activeWorksheet.get_Range("T" + i).Value2 = sb.ToString();
                    }
                    else
                    {
                        activeWorksheet.get_Range("T" + i).Value2 = "";
                    }
                    //else
                    //{
                    //    if (activeWorksheet.get_Range("N" + i).DisplayFormat.Font.Bold == true)
                    //    {
                    //        GetDaTouBi(sendcity, sendaddress, receivercity, receiveraddress, sb);
                    //        activeWorksheet.get_Range("T" + i).Value2 = sb.ToString();
                    //    }
                    //    else
                    //    {
                    //        activeWorksheet.get_Range("T" + i).Value2 = "";
                    //    }
                    //}
                    if (paixu == "是")
                    {
                        string namesort = activeWorksheet.get_Range("P" + i).Value2.ToString();
                        int time = 0;
                        foreach (var item in name)
                        {
                            if (item.Contains(namesort)) time++;
                        }                      
                        if (time < 2)
                        {
                         
                            activeWorksheet.get_Range("U" + i).Value2 = 0;                           
                            time2++;

                        }
                        else
                        {
                            activeWorksheet.get_Range("U" + i).Value2 = 1;
                          
                        }                       

                      
                    }
                    Application.StatusBar = "正在检验地址并生成大头笔.....已完成（" + i + "/" + rowNumber + ")";

                }             
               
            }
            if (paixu == "是") {
            activeWorksheet.get_Range("A2:U" + rowNumber).Sort(activeWorksheet.get_Range("U2:U" + rowNumber), XlSortOrder.xlAscending,
                                        missing, activeWorksheet.get_Range("P2:P" + rowNumber), XlSortOrder.xlAscending, activeWorksheet.get_Range("O2:O" + rowNumber), XlSortOrder.xlAscending, XlYesNoGuess.xlNo, missing, missing, XlSortOrientation.xlSortColumns, XlSortMethod.xlPinYin, XlSortDataOption.xlSortNormal, XlSortDataOption.xlSortNormal, XlSortDataOption.xlSortNormal);
            activeWorksheet.Columns[21].Delete();
            activeWorksheet.get_Range("A2:T" + time2).Sort(activeWorksheet.get_Range("B2:B" + time2), XlSortOrder.xlAscending,
                                  missing, activeWorksheet.get_Range("C2:C" + time2), XlSortOrder.xlAscending, missing, XlSortOrder.xlAscending, XlYesNoGuess.xlNo, missing, missing, XlSortOrientation.xlSortColumns, XlSortMethod.xlPinYin, XlSortDataOption.xlSortNormal, XlSortDataOption.xlSortNormal, XlSortDataOption.xlSortNormal);
            }
            stopwatch.Stop();
            TimeSpan timeSpan = stopwatch.Elapsed;
            double seconds = timeSpan.TotalSeconds;
            Application.StatusBar = "完成\\（o.o）/......耗时"+seconds.ToString("#0.0")+"秒";
         
           

        }
        #endregion
        public string GetDaTouBi(string sendcity, string sendaddress, string receivercity, string receiveraddress, StringBuilder sb)
        {
            sb.Append(PostDateFunc.GetRemaike2(sendcity, sendaddress, receivercity, receiveraddress));
            return sb.ToString();
        }
   
     
        public class IndexerClass
        {
            int rownum = Globals.Sheet1.UsedRange.Rows.Count;
            private string[] name = new string[Globals.Sheet1.UsedRange.Rows.Count];

            //索引器必须以this关键字定义，其实这个this就是类实例化之后的对象
            public string this[int index]
            {
                //实现索引器的get方法
                get
                {
                    if (index < 2)
                    {
                        return name[index];
                    }
                    return null;
                }

                //实现索引器的set方法
                set
                {
                    if (index < rownum)
                    {
                        name[index] = value;
                    }
                }
            }

           
        }
        IWorkbook workbook;
        #region  
        public System.Data.DataTable ImportExcelFile(string filePath,int sheetnum,int hrow,int srow)
        {
            #region
            
            try
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    //file.Position = 0;
                    if (filePath.EndsWith(".xls"))
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                    else if (filePath.EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(file);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            #endregion

            ISheet sheet = workbook.GetSheetAt(sheetnum);
            System.Data.DataTable drtable = new System.Data.DataTable();
            IRow headerRow = sheet.GetRow(hrow);
            int cellCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                drtable.Columns.Add(column);
            }

            for (int i = (sheet.FirstRowNum + srow); i <= rowCount; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = drtable.NewRow();

                if (row != null)
                {
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            dataRow[j] = row.GetCell(j);
                        }
                       
                    }
                }

                drtable.Rows.Add(dataRow);
            }
            return drtable;
        }
        #endregion
        public System.Data.DataTable samenum(int rowCount,int colCount, System.Data.DataTable datable)
        {
            System.Data.DataTable drtable = new System.Data.DataTable();
            string bnum;
            var selection = Globals.ThisWorkbook.Application.Selection;          
                DataColumn column = new DataColumn("行号");
                drtable.Columns.Add(column);
                column = new DataColumn("地址");
                drtable.Columns.Add(column);
            if (colCount == 3)
            {
                column = new DataColumn("优先级");
                drtable.Columns.Add(column);
            }
            
            for (int i = 1; i <= rowCount-1; i++)
                {
                    int num = selection.Rows[i].Row;
                    DataRow dataRow = drtable.NewRow();              
                    dataRow[0] =num;
                    dataRow[1] = Globals.Sheet1.get_Range("N" + num).Value2;
                if (colCount == 3)
                {
                    bnum = Globals.Sheet1.get_Range("D" + num).Value2.ToString();
                    dataRow[2] = GetStrName(datable, "%" + bnum + "%", "类型", " like", "优先级");
                }
                drtable.Rows.Add(dataRow);
                }

                    
                

                
           
            return drtable;
        }
        private string GetStrName(System.Data.DataTable dtable, string Name,string Keyname,string type,string Pname)
        {
           
            DataRow[] dr = dtable.Select(Keyname+type+"'" + Name+"'");
            if (dr.Length > 0)
            {
                Name = dr[0][Pname].ToString();
            }
            else
            {
                Name = null;
            }
            return Name;
        }

        public static decimal GetNumber(string str)
        {
            decimal result = 0;
            if (str != null && str != string.Empty)
            {            
                str = Regex.Replace(str, @"[^\d.\d]", "");                
                if (Regex.IsMatch(str, @"^[+-]?\d*[.]?\d*$"))
                {
                    result = decimal.Parse(str);
                }
            }
            return result;
        }
        public static string Getcn(string str)
        {
            string result="";
            if (str != null && str != string.Empty)
            {
                result = Regex.Replace(str, @"[A-Za-z0-9]", "");            
            }
            return result;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            var selection = Globals.ThisWorkbook.Application.Selection;            
            int a = selection.Rows.Count+1;
            Globals.Sheet1.get_Range("U"+ selection.Rows[1].Row+":AA"+ selection.Rows[selection.Rows.Count].Row).Value2="";
            string cpqd = Globals.Sheet4.get_Range("B8").Value2;
            string yfjs = Globals.Sheet4.get_Range("B10").Value2;
            string tsqy = Globals.Sheet4.get_Range("D10").Value2;
            string kyqy = Globals.Sheet4.get_Range("D11").Value2;
            string emsqy = Globals.Sheet4.get_Range("D12").Value2;
            string emsqyc = Globals.Sheet4.get_Range("E12").Value2;
            System.Data.DataTable drtable =ImportExcelFile(cpqd,0,1,2);
            //Excel.Workbook wb =Application.Workbooks.Open(cpqd, missing, true, missing, missing, missing,missing, missing, missing, true, missing, missing, missing, missing, missing);
            //Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);
            //Globals.Sheet1.get_Range("O2").Value2= ws.get_Range("B12").Value2;
            System.Data.DataTable sametable=samenum(a,2,null);
            for (int i = 1; i < a; i++)
            {
                if (selection.Cells.SpecialCells(XlCellType.xlCellTypeVisible, 12).Rows[i].Hidden == false)
                {
                    int num = selection.Rows[i].Row;
                    int snum = Convert.ToInt32(Globals.Sheet1.get_Range("C" + num).Value2);
                    if (Globals.Sheet1.get_Range("L" + num).Value2 == null)
                    {
                        continue;
                    }
                    string bnum = Globals.Sheet1.get_Range("D" + num).Value2.ToString();
                    string pname = Globals.Sheet1.get_Range("B" + num).Value2.ToString();
                    string province = Globals.Sheet1.get_Range("K" + num).Value2;
                    string city = Globals.Sheet1.get_Range("L" + num).Value2;
                    string xian = Globals.Sheet1.get_Range("M" + num).Value2;
                    string address = Globals.Sheet1.get_Range("N" + num).Value2;
                    string size = GetStrName(drtable, bnum, "编码", "=", "规格");
                    string xiangzhong = GetStrName(drtable, bnum, "编码", "=", "每箱重量（kg）");
                    string jianzhong = GetStrName(drtable, bnum, "编码", "=", "单件重量（kg）");
                    string sfjzx = GetStrName(drtable, bnum, "编码", "=", "顺丰计重箱");
                    string sfjzt = GetStrName(drtable, bnum, "编码", "=", "顺丰计重台");
                    string kyjzx = GetStrName(drtable, bnum, "编码", "=", "跨越物流");
                    string hanghao = GetStrName(sametable, address, "地址", "=", "行号");
                    string xs = "";
                    decimal sizenum = GetNumber(size);
                    string danwei = Getcn(size);
                    if (sizenum != 0)
                    {
                        decimal xiangshu = snum / Convert.ToInt32(sizenum);
                        decimal jianshu = snum % sizenum;
                        if (jianshu != 0 && xiangshu != 0)
                        {
                            xs = xiangshu + "箱" + pname + "+" + jianshu + danwei.Substring(0, 1) + pname;
                        }
                        else if (jianshu != 0 && xiangshu == 0)
                        {
                            xs = jianshu + danwei.Substring(0, 1) + pname;
                        }
                        else
                        {
                            xs = xiangshu + "箱" + pname;
                        }
                        Globals.Sheet1.get_Range("U" + num).Value2 = xs;
                        if (GetNumber(sfjzx) > 0)
                        { xiangzhong = GetNumber(sfjzx).ToString(); }
                        if (GetNumber(sfjzt) > 0)
                        { jianzhong = GetNumber(sfjzt).ToString(); }
                        decimal zhongliangb = xiangshu * Convert.ToDecimal((xiangzhong != "") ? xiangzhong : "0");
                        decimal zhongliangs = jianshu * Convert.ToDecimal((jianzhong != "") ? jianzhong : "0");
                        decimal zhongliangt = zhongliangb + zhongliangs;//重量
                        decimal hebingzhong = Convert.ToDecimal(Globals.Sheet1.get_Range("V" + hanghao).Value2);
                        decimal ZL = hebingzhong + zhongliangt;
                        Globals.Sheet1.get_Range("V" + hanghao).Value2 = ZL;
                        //计算顺丰邮费
                        System.Data.DataTable SFtable = ImportExcelFile(yfjs, 0, 2, 3);
                        string yunfei = "0";
                        //if(tsqy.Contains(city.Substring(0,2)))
                        //{
                        //    city= province;
                        //}
                        string minpay = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "最低消费(顺丰)");
                        if (minpay == null) { minpay = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "最低消费(顺丰)"); }
                        string qizhong = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "发货起重(顺丰)");
                        if (qizhong == null) { qizhong = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "发货起重(顺丰)"); }
                        qizhong = qizhong == "" ? "0 " : qizhong;
                        minpay = minpay == "" ? "0" : minpay;
                        if (ZL > Convert.ToInt32(qizhong))
                        {
                            yunfei = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "顺丰物流");
                            if (yunfei == "0")
                            {
                                yunfei = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "顺丰物流");
                            }
                        }


                        decimal jiage = ZL * Convert.ToDecimal(yunfei);
                        decimal finaljiage;
                        //if (yunfei != "0")
                        //{
                        finaljiage = jiage > Convert.ToDecimal(minpay) ? jiage : Convert.ToDecimal(minpay);
                        //}
                        //else
                        //{
                        //    finaljiage = 0;
                        //}
                        Globals.Sheet1.get_Range("W" + hanghao).Value2 = finaljiage;
                        //计算跨越邮费
                        //System.Data.DataTable KYtable = ImportExcelFile(yfjs, 0,2,3);
                        city = Globals.Sheet1.get_Range("L" + num).Value2;
                        //if (kyqy.Contains(province.Substring(0, 2)) || kyqy.Contains(city.Substring(0, 2)))
                        //{
                        if (GetNumber(kyjzx) > 0)
                        {
                            xiangzhong = GetNumber(kyjzx).ToString();
                        }
                        if (GetNumber(sfjzt) > 0)
                        {
                            jianzhong = GetNumber(sfjzt).ToString();
                        }
                        zhongliangb = xiangshu * Convert.ToDecimal((xiangzhong != "") ? xiangzhong : "0");
                        zhongliangs = jianshu * Convert.ToDecimal((jianzhong != "") ? jianzhong : "0");
                        zhongliangt = zhongliangb + zhongliangs;
                        hebingzhong = Convert.ToDecimal(Globals.Sheet1.get_Range("X" + hanghao).Value2);
                        ZL = hebingzhong + zhongliangt;
                        qizhong = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "发货起重(跨越)");
                        if (qizhong == null) { qizhong = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "发货起重(跨越)"); }
                        yunfei = "0";
                        minpay = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "最低消费(跨越)");
                        if (minpay == null) { minpay = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "最低消费(跨越)"); }
                        qizhong = qizhong == "" ? "0 " : qizhong;
                        minpay = minpay == "" ? "0" : minpay;
                        if (ZL <= 100 & ZL >= Convert.ToInt32(qizhong))
                        {
                            yunfei = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "≤100");
                            if (yunfei == null)
                            {
                                yunfei = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "≤100");
                            }

                        }
                        else if (ZL > 100 && ZL <= 300)
                        {
                            yunfei = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "101-300");
                            if (yunfei == null)
                            {
                                yunfei = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "101-300");
                            }
                        }
                        else if (ZL > 300 && ZL <= 500)
                        {
                            yunfei = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "301-500");
                            if (yunfei == null)
                            {
                                yunfei = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "301-500");
                            }
                        }
                        else if (ZL > 500 && ZL <= 1000)
                        {
                            yunfei = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "501-1000G");
                            if (yunfei == null)
                            {
                                yunfei = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "501-1000");
                            }
                        }
                        else if (ZL > 1000)
                        {
                            yunfei = GetStrName(SFtable, "%" + city.Substring(0, 2) + "%", "目的城市", " like", "＞1000");
                            if (yunfei == null)
                            {
                                yunfei = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "目的城市", " like", "＞1000");
                            }
                        }
                        yunfei = yunfei == "" ? "0" : yunfei;
                        jiage = ZL * Convert.ToDecimal(yunfei);
                        //if (yunfei != "0")
                        //{
                        finaljiage = jiage > Convert.ToDecimal(minpay) ? jiage : Convert.ToDecimal(minpay);
                        //}
                        //else
                        //{
                        //    finaljiage = 0;
                        //}
                        Globals.Sheet1.get_Range("X" + hanghao).Value2 = ZL;
                        Globals.Sheet1.get_Range("Y" + hanghao).Value2 = finaljiage;
                        //}
                        //计算ems邮费 
                        //if (emsqy.Contains(province.Substring(0, 2)) && !emsqyc.Contains(xian.Substring(0, 2)))
                        //{

                        //    string emssz = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "地点", " like", "EMS首重1");
                        //    string emssz2 = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "地点", " like", "EMS首重2");
                        //    string emsxz = GetStrName(SFtable, "%" + province.Substring(0, 2) + "%", "地点", " like", "EMS续重");
                        //    xiangzhong = GetStrName(drtable, bnum, "编码", "=", "每箱重量（kg）");
                        //    jianzhong = GetStrName(drtable, bnum, "编码", "=", "单件重量（kg）");
                        //    zhongliangb = xiangshu * Convert.ToDecimal((xiangzhong != "") ? xiangzhong : "0");
                        //    zhongliangs = jianshu * Convert.ToDecimal((jianzhong != "") ? jianzhong : "0");
                        //    zhongliangt = zhongliangb + zhongliangs;
                        //    double syzl = Convert.ToDouble(zhongliangt);

                        //    if (syzl > 0.5)
                        //    {
                        //        syzl = syzl - 1;
                        //        int time = 0;
                        //        for (int k=0; syzl > 0.5; k++)
                        //        {
                        //            time++;
                        //            syzl = syzl - 0.5;       

                        //        }
                        //        if (syzl > 0 && syzl < 0.5) { time++; }
                        //         jiage = Convert.ToInt32(emssz2) + 1 +(time* Convert.ToInt32(emsxz));
                        //    }
                        //    else
                        //    {
                        //        jiage = Convert.ToInt32(emssz) + 1;
                        //    }
                        //    Globals.Sheet1.get_Range("Z" + num).Value2 = jiage + "元";
                        //}
                    }

                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var selection = Globals.ThisWorkbook.Application.Selection;
            int a = selection.Rows.Count + 1;
            Globals.Sheet1.get_Range("H" + selection.Rows[1].Row + ":H" + selection.Rows[selection.Rows.Count].Row).Value2 = "";
            string pydq = Globals.Sheet4.get_Range("D13").Value2;
            string yfjs = Globals.Sheet4.get_Range("B10").Value2;
            string cpqd = Globals.Sheet4.get_Range("B8").Value2;
            System.Data.DataTable yftable = ImportExcelFile(yfjs,1,1,2);
            System.Data.DataTable emstable = ImportExcelFile(yfjs, 2, 3, 4);
            System.Data.DataTable drtable = ImportExcelFile(cpqd, 0, 1, 2);
            string emsqy = Globals.Sheet4.get_Range("D12").Value2;
            string emsqyc = Globals.Sheet4.get_Range("E12").Value2;
            string yunfei="0";
            string yunfeix = "0";
            System.Data.DataTable sametable = samenum(a,3,yftable);
            for (int i = 1; i < a; i++)
            {
                if (selection.Cells.SpecialCells(XlCellType.xlCellTypeVisible, 12).Rows[i].Hidden == false)
                {
                    int num = selection.Rows[i].Row;
                    if (Globals.Sheet1.get_Range("L" + num).Value2 == null)
                    {
                        continue;
                    }
                    int snum = Convert.ToInt32(Globals.Sheet1.get_Range("C" + num).Value2);
                    string bnum = Globals.Sheet1.get_Range("D" + num).Value2.ToString();
                    string shuliang = Globals.Sheet1.get_Range("C" + num).Value2.ToString();
                    string province = Globals.Sheet1.get_Range("K" + num).Value2 == null ? "  " : Globals.Sheet1.get_Range("K" + num).Value2;
                    string city = Globals.Sheet1.get_Range("L" + num).Value2;
                    string xian = Globals.Sheet1.get_Range("M" + num).Value2;
                    string address = Globals.Sheet1.get_Range("N" + num).Value2;
                    string size = GetStrName(drtable, bnum, "编码", "=", "规格");
                    object maxnum = sametable.Compute("Min(优先级)", "地址='" + address + "'");
                    string hanghao = GetStrName(sametable, address, "优先级='" + maxnum + "'and 地址", "=", "行号");
                    int byf = 0;
                    int jiage = 0;
                    if (emsqy.Contains(province.Substring(0, 2)) && !emsqyc.Contains(xian.Substring(0, 2)))
                    {
                        decimal sizenum = GetNumber(size);
                        if (sizenum != 0)
                        {
                            decimal xiangshu = snum / Convert.ToInt32(sizenum);
                            decimal jianshu = snum % sizenum;
                            string emssz = GetStrName(emstable, "%" + province.Substring(0, 2) + "%", "地点", " like", "EMS首重1");
                            string emssz2 = GetStrName(emstable, "%" + province.Substring(0, 2) + "%", "地点", " like", "EMS首重2");
                            string emsxz = GetStrName(emstable, "%" + province.Substring(0, 2) + "%", "地点", " like", "EMS续重");
                            string xiangzhong = GetStrName(drtable, bnum, "编码", "=", "每箱重量（kg）");
                            string jianzhong = GetStrName(drtable, bnum, "编码", "=", "单件重量（kg）");
                            decimal zhongliangb = xiangshu * Convert.ToDecimal((xiangzhong != "") ? xiangzhong : "0");
                            decimal zhongliangs = jianshu * Convert.ToDecimal((jianzhong != "") ? jianzhong : "0");
                            decimal zhongliangt = zhongliangb + zhongliangs;
                            double syzl = Convert.ToDouble(zhongliangt);
                            if (syzl > 0.5)
                            {
                                syzl = syzl - 1;
                                int time = 0;
                                for (int k = 0; syzl > 0.5; k++)
                                {
                                    time++;
                                    syzl = syzl - 0.5;

                                }
                                if (syzl > 0 && syzl < 0.5) { time++; }
                                jiage = Convert.ToInt32(emssz2) + 1 + (time * Convert.ToInt32(emsxz));
                            }
                            else
                            {
                                jiage = Convert.ToInt32(emssz) + 1;
                            }

                            byf = jiage;
                        }
                    }
                    else
                    {
                        if (pydq.Contains(province.Substring(0, 2)))
                        {
                            if (num == Convert.ToInt32(hanghao))
                            {
                                yunfei = GetStrName(yftable, "%" + bnum + "%", "类型", " like", "偏远首件");
                                yunfeix = GetStrName(yftable, "%" + bnum + "%", "类型", " like", "偏远续件");
                            }
                            else
                            {
                                yunfei = GetStrName(yftable, "%" + bnum + "%", "类型", " like", "偏远续件");
                            }

                        }
                        else if (province.Substring(0, 2) == "广东")
                        {
                            yunfei = GetStrName(yftable, "%" + bnum + "%", "类型", " like", "省内");
                            yunfeix = yunfei;
                        }
                        else if (province.Substring(0, 2) == "香港" || province.Substring(0, 2) == "台湾")
                        {
                            yunfei = GetStrName(yftable, "%" + bnum + "%", "类型", " like", "首件");
                        }
                        else
                        {
                            if (num == Convert.ToInt32(hanghao))
                            {
                                yunfei = GetStrName(yftable, "%" + bnum + "%", "类型", " like", "省外首件");
                                yunfeix = GetStrName(yftable, "%" + bnum + "%", "类型", " like", "省外续件");
                            }
                            else
                            {
                                yunfei = GetStrName(yftable, "%" + bnum + "%", "类型", " like", "省外续件");
                            }
                        }
                        byf = (Convert.ToInt32(yunfei) + Convert.ToInt32(yunfeix) * (Convert.ToInt32(shuliang) - 1)) + Convert.ToInt32(Globals.Sheet1.get_Range("H" + hanghao).Value2);

                    }
                    Globals.Sheet1.get_Range("H" + hanghao).Value2 = byf;

                }
            }

        }

       public static void test(){ string test = ""; }
    }
}
